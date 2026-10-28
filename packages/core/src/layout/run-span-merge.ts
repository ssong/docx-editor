// One span per run segment (`SemanticLayoutOptions.spans: 'run'`).
//
// The line breaker places Latin text one word at a time, so a laid-out line holds one span per
// word. Once the line is closed and aligned, a seam after a word's space carries no
// information when nothing stretches it: the two spans share every decoration and abut. This
// module merges such seams, the way `coalesceIdeographicSpans` merges ideographic ones. A
// merged span's width is the sum of the advances the line was broken against; nothing is
// measured again.
//
// Two places merge, so a line is merged where the merge does not cost a copy:
//
// - At line close (`mergeClosedLineSpans`), before the paragraph cache stores the break, for
//   lines that alignment publishes unchanged: left-aligned lines and the last line of a
//   justified paragraph. Those spans are shared between the cache and the laid-out page, so the
//   merge shrinks both.
// - After alignment (`mergeAlignedSpans`), for spans that alignment copied anyway: centred and
//   right-aligned lines, and lines placed at a pen offset. Justified seams keep their spans,
//   because alignment puts slack in them.
//
// The option is a layout INPUT for the whole pass, installed around it like the field results
// mode, and folded into the producer (`field-results-producer.ts`) so a cache never answers one
// mode with the other's break.

import { PAGE_BREAK_CHAR } from '@docx-editor.dev/core/store';
import { automaticTextSpacingBefore } from './cjk-automatic-spacing.ts';
import { withBox } from './compact-box.ts';
import type { Alignment } from './paragraph-alignment.ts';
import { decorationsMatch, type PendingLine } from './pending-line.ts';
import type { StyleSpanRecord } from './semantic-records.ts';

/**
 * How many spans a laid-out line holds for its text.
 *
 * - `word` (the default): one span per word, as the line breaker places them.
 * - `run`: one span per run segment. Adjacent words of one run join into one span when no
 *   justification stretches the space between them. Text, model ranges, line and page breaks,
 *   and every glyph position stay the same; only the span boundaries move.
 * @beta
 */
export type SpanGranularity = 'word' | 'run';

let current: SpanGranularity = 'word';

/** The span granularity of the layout pass in progress; `word` outside any pass. */
export function currentSpanGranularity(): SpanGranularity {
  return current;
}

/**
 * Run `run` with `mode` installed for the layout it does, then restore the outer mode, also
 * when `run` throws. `undefined` keeps the outer mode.
 */
export function withSpanGranularity<T>(mode: SpanGranularity | undefined, run: () => T): T {
  if (mode === undefined || mode === current) return run();
  const outer = current;
  current = mode;
  try {
    return run();
  } finally {
    current = outer;
  }
}

/** Characters whose advance or offset a merge would hide from the code that reads them. */
const SEAM_UNSAFE = /[\t\n\v\f\r\u00ad\u200b-\u200f\u2028\u2029\u2060\ufeff]/u;

/**
 * Whether a span is ordinary text a merge may join: one model offset per drawn UTF-16 unit, no
 * layout-owned advance, no bidi shaping, and nothing that is measured against the span's own
 * text (`caretEdges`, `glyphOffsetPt`).
 */
function plainText(span: StyleSpanRecord): boolean {
  return (
    span.text.length > 0 &&
    span.text.length === span.range.end - span.range.start &&
    span.text !== PAGE_BREAK_CHAR &&
    !SEAM_UNSAFE.test(span.text) &&
    span.style.shaping === undefined &&
    span.caretEdges === undefined &&
    !span.glyphOffsetPt &&
    span.tabLeader === undefined &&
    !span.lineEndWhitespace &&
    !span.optionalHyphenBreak &&
    !span.projected &&
    !span.equation &&
    !span.unsupported &&
    !span.noteSeparator &&
    !span.noteNav
  );
}

/** Two spans whose x positions are this close abut. */
const SEAM_TOLERANCE_PT = 0.001;

/**
 * Whether the seam between two adjacent spans of one line carries no information: the first
 * ends a word with its space, both are plain text of one paragraph that follow each other in
 * the model, they abut, and they share every decoration by reference.
 */
function seamJoins(previous: StyleSpanRecord, next: StyleSpanRecord): boolean {
  if (!previous.text.endsWith(' ')) return false;
  if (!plainText(previous) || !plainText(next)) return false;
  if (previous.range.paragraphId !== next.range.paragraphId) return false;
  if (previous.range.end !== next.range.start) return false;
  if (next.wrapAdvanceBefore !== undefined || automaticTextSpacingBefore(next) !== 0) return false;
  if (Math.abs(previous.box.x + previous.box.width - next.box.x) > SEAM_TOLERANCE_PT) return false;
  if (previous.box.y !== next.box.y || previous.box.height !== next.box.height) return false;
  return decorationsMatch(previous, next);
}

/**
 * `spans` with every joinable seam merged, or `spans` itself when no seam joins.
 * `joinable` can refuse a seam the caller must keep.
 */
function mergeSeams(
  spans: readonly StyleSpanRecord[],
  joinable: (previous: StyleSpanRecord, next: StyleSpanRecord) => boolean
): readonly StyleSpanRecord[] {
  let merged: StyleSpanRecord[] | undefined;
  for (let index = 0; index < spans.length; ) {
    let end = index;
    while (
      end + 1 < spans.length &&
      seamJoins(spans[end]!, spans[end + 1]!) &&
      joinable(spans[end]!, spans[end + 1]!)
    ) {
      end += 1;
    }
    const first = spans[index]!;
    if (end === index) {
      merged?.push(first);
      index += 1;
      continue;
    }
    merged ??= spans.slice(0, index);
    const last = spans[end]!;
    let text = '';
    for (let cursor = index; cursor <= end; cursor += 1) text += spans[cursor]!.text;
    merged.push({
      ...first,
      range: { ...first.range, end: last.range.end },
      text,
      // The joined spans abut, so the far edge of the last is the sum of their advances.
      box: withBox(first.box, { width: last.box.x + last.box.width - first.box.x }),
    });
    index = end + 1;
  }
  return merged ?? spans;
}

/**
 * Whether alignment publishes this broken line's spans unchanged, so a merge before the cache
 * stores the break is safe: a left-aligned line, or a justified paragraph's last line that
 * borrowed no space. Right-to-left paragraphs align and reorder their spans, so never.
 */
function alignsUnchanged(
  alignment: Alignment,
  paragraphRtl: boolean,
  line: PendingLine,
  isLastLine: boolean
): boolean {
  if (paragraphRtl) return false;
  if (alignment === 'left') return true;
  return alignment === 'both' && isLastLine && line.spaceShrink !== true;
}

/**
 * Merge the run seams of a paragraph's freshly broken lines in place, before the paragraph
 * cache stores them. Only lines that alignment publishes unchanged merge here (see the module
 * comment); the others merge after alignment, in {@link mergeAlignedSpans}.
 */
export function mergeClosedLineSpans(
  lines: readonly PendingLine[],
  alignment: Alignment | undefined,
  paragraphRtl: boolean
): void {
  if (current !== 'run' || alignment === undefined) return;
  for (const [index, line] of lines.entries()) {
    if (line.spans.length < 2 || Object.isFrozen(line.spans)) continue;
    if (!alignsUnchanged(alignment, paragraphRtl, line, index === lines.length - 1)) continue;
    const merged = mergeSeams(line.spans, () => true);
    if (merged === line.spans) continue;
    line.spans.length = 0;
    for (const span of merged) line.spans.push(span);
  }
}

/**
 * An aligned line's spans with their run seams merged, in a `run` pass.
 *
 * Only seams between spans that alignment or placement copied merge. A frozen span is shared
 * with the paragraph cache, and merging it would publish a second copy beside the one the cache
 * keeps. Seams that alignment stretched do not abut, so justified spans stay apart.
 */
export function mergeAlignedSpans(spans: readonly StyleSpanRecord[]): readonly StyleSpanRecord[] {
  if (current !== 'run' || spans.length < 2) return spans;
  return mergeSeams(
    spans,
    (previous, next) => !Object.isFrozen(previous) && !Object.isFrozen(next)
  );
}
