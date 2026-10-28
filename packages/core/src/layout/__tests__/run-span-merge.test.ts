// One span per run segment (`spans: 'run'`) changes span boundaries and nothing else.
//
// Every test lays the same document out twice, one span per word and one per run segment, and
// asks both layouts the questions an editor asks: where each line breaks, where the caret for
// every offset is, which offset a click lands on, which bands a selection, a find hit or a
// comment paints, which word a double-click selects, and what text a line holds. The answers
// must match. Only the number of spans may differ.

import { expect, test } from 'bun:test';
import { readOoxmlPart } from '../../store/index.ts';
import { everyStoryOrder } from '../document-order.ts';
import { isCompactBox } from '../compact-box.ts';
import { createParagraphLayoutCache } from '../layout-cache.ts';
import type { PendingLine } from '../pending-line.ts';
import type { SpanGranularity } from '../run-span-merge.ts';
import { hitTestPage } from '../semantic-hit-test.ts';
import { caretAt, paragraphTextFromLayout, spansInSelection } from '../semantic-interaction.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
import type { LineRecord, SemanticLayout, StyleSpanRecord } from '../semantic-records.ts';
import { linesOf } from '../semantic-records.ts';
import { keyedRangeRects, selectionRects } from '../selection-rects.ts';
import { wordRangeAt } from '../../editor/surface-selection-ops.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
/** 6pt per character for the 11pt runs below, so a 150pt measure holds 25 characters. */
const measurer = createFixedMeasurer(6, 14);
const WIDTH = 150;

const run = (text: string, rPr = '') =>
  `<w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (runs: string, pPr = '') => `<w:p><w:pPr>${pPr}</w:pPr>${runs}</w:p>`;

const PROSE = 'agreement balance contract delivery estimate forecast guarantee heading';

/** One paragraph per case, in this order. */
const CASES = {
  left: paragraph(run(`${PROSE} `) + run('invoice journal ', '<w:b/>') + run('ledger margin.')),
  center: paragraph(run(PROSE), '<w:jc w:val="center"/>'),
  right: paragraph(run(PROSE), '<w:jc w:val="right"/>'),
  justified: paragraph(run(`${PROSE} invoice journal ledger.`), '<w:jc w:val="both"/>'),
  rtl: paragraph(
    run('שלום עולם ', '<w:rtl/>') + run('hello world ') + run('אבג דהו זחט יכל', '<w:rtl/>'),
    '<w:bidi/>'
  ),
  mixedDirection: paragraph(run('alpha beta ') + run('אבג דהו', '<w:rtl/>') + run(' gamma delta')),
  cjk: paragraph(run('天地玄黄 alpha beta 宇宙洪荒日月盈昃 gamma delta 辰宿列张')),
  revision:
    `<w:p>${run('alpha beta gamma ')}` +
    `<w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">${run('delta epsilon ')}</w:ins>` +
    `${run('zeta eta theta iota kappa')}</w:p>`,
  tabs: paragraph(run('one two') + '<w:r><w:tab/></w:r>' + run('three four five six seven')),
} as const;

const BODY = Object.values(CASES).join('');
const CASE_INDEX = Object.fromEntries(Object.keys(CASES).map((name, index) => [name, index]));

function part() {
  const parsed = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${BODY}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/xml',
  });
  if (!parsed.ok) throw new Error(parsed.reason);
  return parsed.part;
}

function layout(
  spans: SpanGranularity,
  cache?: ReturnType<typeof createParagraphLayoutCache<readonly PendingLine[]>>
): SemanticLayout {
  return layoutSemanticDocument(part(), 0, {
    measurer,
    spans,
    producer: 'test',
    ...(cache ? { cache } : {}),
    geometry: { width: WIDTH, height: 2000, margin: { top: 0, bottom: 0, left: 0, right: 0 } },
  });
}

const word = layout('word');
const merged = layout('run');
const order = everyStoryOrder(word);

function linesFor(doc: SemanticLayout, name: keyof typeof CASES): LineRecord[] {
  const id = order[CASE_INDEX[name]!]!;
  return linesOf(doc).filter((line) => line.range.paragraphId === id);
}

const spanCount = (lines: readonly LineRecord[]) =>
  lines.reduce((sum, line) => sum + line.spans.length, 0);

/** Horizontal intervals each row covers, with touching rectangles joined. */
function coverage(rects: readonly { x: number; y: number; width: number; height: number }[]) {
  const rows = new Map<string, [number, number][]>();
  for (const rect of rects) {
    if (rect.width <= 0) continue;
    const key = `${rect.y.toFixed(3)}:${rect.height.toFixed(3)}`;
    const row = rows.get(key) ?? [];
    row.push([rect.x, rect.x + rect.width]);
    rows.set(key, row);
  }
  const result: Record<string, string[]> = {};
  for (const [key, row] of [...rows].sort(([a], [b]) => a.localeCompare(b))) {
    row.sort((a, b) => a[0] - b[0]);
    const joined: [number, number][] = [];
    for (const [from, to] of row) {
      const last = joined[joined.length - 1];
      if (last && from <= last[1] + 1e-6) last[1] = Math.max(last[1], to);
      else joined.push([from, to]);
    }
    result[key] = joined.map(([from, to]) => `${from.toFixed(4)}..${to.toFixed(4)}`);
  }
  return result;
}

test('every line breaks at the same place and holds the same text', () => {
  const describe = (doc: SemanticLayout) =>
    linesOf(doc).map((line) => ({
      range: line.range,
      contentX: line.contentX,
      box: line.box,
      text: line.spans.map((span) => span.text).join(''),
    }));
  expect(describe(merged)).toEqual(describe(word));
  for (const id of order) {
    expect(paragraphTextFromLayout(merged, id)).toBe(paragraphTextFromLayout(word, id));
  }
});

test('left, centred and right-aligned lines hold one span per run segment', () => {
  // The left case has three runs; its wrapped lines each hold at most one span per run.
  for (const line of linesFor(merged, 'left')) {
    expect(line.spans.length).toBeLessThanOrEqual(3);
  }
  for (const name of ['center', 'right'] as const) {
    for (const line of linesFor(merged, name)) expect(line.spans).toHaveLength(1);
  }
  expect(spanCount(linesFor(merged, 'left'))).toBeLessThan(spanCount(linesFor(word, 'left')));
});

test('a merged span starts where its first word started and ends where its last word ended', () => {
  for (const name of Object.keys(CASES) as (keyof typeof CASES)[]) {
    const before = linesFor(word, name);
    const after = linesFor(merged, name);
    for (const [index, line] of after.entries()) {
      const words = [...before[index]!.spans];
      for (const span of line.spans) {
        const first = words[0]!;
        let text = '';
        let last = first;
        while (text.length < span.text.length) {
          last = words.shift()!;
          text += last.text;
        }
        expect(text).toBe(span.text);
        expect(span.range.start).toBe(first.range.start);
        expect(span.range.end).toBe(last.range.end);
        expect(span.box.x).toBeCloseTo(first.box.x, 9);
        expect(span.box.x + span.box.width).toBeCloseTo(last.box.x + last.box.width, 9);
        expect(span.box.y).toBe(first.box.y);
        expect(span.box.height).toBe(first.box.height);
        expect(span.style).toEqual(first.style);
      }
    }
  }
});

test('justified lines keep a span boundary wherever justification stretches the space', () => {
  const before = linesFor(word, 'justified');
  const after = linesFor(merged, 'justified');
  expect(after.length).toBeGreaterThan(2);
  let stretched = 0;
  for (const [index, line] of before.entries()) {
    const starts = new Map(after[index]!.spans.map((span) => [span.range.start, span.box.x]));
    for (const [spanIndex, span] of line.spans.entries()) {
      const previous = line.spans[spanIndex - 1];
      if (!previous || span.box.x - (previous.box.x + previous.box.width) <= 0.001) continue;
      // The slack paint draws as word spacing stays between two spans, at the same place.
      stretched += 1;
      expect(starts.get(span.range.start)).toBe(span.box.x);
    }
  }
  expect(stretched).toBeGreaterThan(0);
  // The last line is set flush left, so its words join.
  expect(after[after.length - 1]!.spans).toHaveLength(1);
  expect(before[before.length - 1]!.spans.length).toBeGreaterThan(1);
});

test('right-to-left and bidi lines keep their spans', () => {
  for (const name of ['rtl', 'mixedDirection'] as const) {
    const before = linesFor(word, name);
    const after = linesFor(merged, name);
    for (const [index, line] of after.entries()) {
      for (const [spanIndex, span] of line.spans.entries()) {
        if (span.style.shaping) expect(span).toEqual(before[index]!.spans[spanIndex]!);
      }
    }
  }
  // A paragraph with right-to-left runs keeps every shaped span apart.
  expect(linesFor(merged, 'rtl').map((line) => line.spans.length)).toEqual(
    linesFor(word, 'rtl').map((line) => line.spans.length)
  );
});

test('CJK text and the Latin words beside it join only at Latin word seams', () => {
  const after = linesFor(merged, 'cjk');
  expect(spanCount(after)).toBeLessThan(spanCount(linesFor(word, 'cjk')));
  for (const line of after) {
    for (const span of line.spans) {
      // A span never runs from a Latin space into an ideograph that automatic spacing moved.
      expect(span.text.length).toBe(span.range.end - span.range.start);
    }
  }
});

test('revision and tab boundaries stay span boundaries', () => {
  const ranges = new Map<number, unknown>();
  for (const line of linesFor(word, 'revision')) {
    for (const span of line.spans) {
      for (let offset = span.range.start; offset < span.range.end; offset += 1)
        ranges.set(offset, span.revisions);
    }
  }
  for (const line of linesFor(merged, 'revision')) {
    for (const span of line.spans) {
      for (let offset = span.range.start; offset < span.range.end; offset += 1)
        expect(ranges.get(offset)).toEqual(span.revisions);
    }
  }
  const tabs = linesFor(merged, 'tabs').flatMap((line) => line.spans);
  expect(tabs.find((span) => span.text.includes('\t'))?.text).toBe('\t');
});

test('the caret for every offset sits at the same place', () => {
  for (const id of order) {
    const length = paragraphTextFromLayout(word, id).length;
    for (let offset = 0; offset <= length; offset += 1) {
      const position = { paragraphId: id, offset };
      const before = caretAt(word, position, measurer);
      const after = caretAt(merged, position, measurer);
      expect({ offset, ...after, lineId: undefined }).toEqual({
        offset,
        ...before,
        lineId: undefined,
      });
    }
  }
});

test('a click anywhere on a line lands on the same offset', () => {
  // Each line's band in page coordinates, from the carets placed on it.
  const bands = new Map<string, { y: number; height: number; left: number; right: number }>();
  for (const id of order) {
    const length = paragraphTextFromLayout(word, id).length;
    for (let offset = 0; offset <= length; offset += 1) {
      const caret = caretAt(word, { paragraphId: id, offset }, measurer)!;
      const band = bands.get(caret.lineId);
      if (!band) {
        bands.set(caret.lineId, {
          y: caret.y,
          height: caret.height,
          left: caret.x,
          right: caret.x,
        });
      } else {
        band.left = Math.min(band.left, caret.x);
        band.right = Math.max(band.right, caret.x);
      }
    }
  }
  for (const band of bands.values()) {
    for (let x = band.left - 3; x <= band.right + 3; x += 1.25) {
      for (const dy of [1, band.height / 2, band.height - 1]) {
        const point = { x, y: band.y + dy };
        const before = hitTestPage(word, 0, point, { measurer });
        const after = hitTestPage(merged, 0, point, { measurer });
        expect({ x, position: after?.position, glyphs: after?.onGlyphs }).toEqual({
          x,
          position: before?.position,
          glyphs: before?.onGlyphs,
        });
      }
    }
  }
});

test('selections, find hits and comment ranges paint the same bands', () => {
  for (const id of order) {
    const length = paragraphTextFromLayout(word, id).length;
    for (let from = 0; from < length; from += 3) {
      for (let to = from + 1; to <= length; to += 5) {
        const selection = {
          anchor: { paragraphId: id, offset: from },
          head: { paragraphId: id, offset: to },
        };
        expect(coverage(selectionRects(merged, selection, order, measurer))).toEqual(
          coverage(selectionRects(word, selection, order, measurer))
        );
        const keyed = [{ key: 'k', from: selection.anchor, to: selection.head }] as const;
        expect(
          coverage(keyedRangeRects(merged, keyed, undefined, measurer).get('k') ?? [])
        ).toEqual(coverage(keyedRangeRects(word, keyed, undefined, measurer).get('k') ?? []));
      }
    }
  }
});

test('a double-click selects the same word, and a selection reports the same formatting', () => {
  const formatting = (spans: readonly StyleSpanRecord[]) => [
    ...new Set(spans.map((span) => span.style)),
  ];
  for (const id of order) {
    const length = paragraphTextFromLayout(word, id).length;
    for (let offset = 0; offset <= length; offset += 1) {
      const position = { paragraphId: id, offset };
      expect(wordRangeAt(merged, position, new Set())).toEqual(
        wordRangeAt(word, position, new Set())
      );
      const selection = { anchor: position, head: { paragraphId: id, offset: length } };
      expect(formatting(spansInSelection(merged, selection, order))).toEqual(
        formatting(spansInSelection(word, selection, order))
      );
    }
  }
});

test('a retained cache shares merged spans with the laid-out lines', () => {
  const cache = createParagraphLayoutCache<readonly PendingLine[]>();
  const first = layout('run', cache);
  const again = layout('run', cache);
  const left = linesFor(first, 'left');
  expect(left[0]!.spans.length).toBeLessThan(linesFor(word, 'left')[0]!.spans.length);
  // The second pass is served from the cache: the same span objects, frozen once.
  for (const [index, line] of linesFor(again, 'left').entries()) {
    expect(line.spans[0]).toBe(left[index]!.spans[0]);
    expect(Object.isFrozen(line.spans[0])).toBe(true);
  }
});

test('one cache never answers one span mode with the other mode’s break', () => {
  const cache = createParagraphLayoutCache<readonly PendingLine[]>();
  const wordFirst = layout('word', cache);
  const runAfter = layout('run', cache);
  const wordAgain = layout('word', cache);
  expect(spanCount(linesFor(wordFirst, 'left'))).toBe(spanCount(linesFor(word, 'left')));
  expect(spanCount(linesFor(runAfter, 'left'))).toBe(spanCount(linesFor(merged, 'left')));
  expect(spanCount(linesFor(wordAgain, 'left'))).toBe(spanCount(linesFor(word, 'left')));
});

test('with compact boxes, a merged span keeps a compact box at the same place', () => {
  // Compare against a retained cache too: its spans are frozen, which decides some merges.
  const exactLayout = layout('run', createParagraphLayoutCache());
  const compact = layout('run', createParagraphLayoutCache({ compactBoxes: true }));
  for (const name of ['left', 'center', 'right', 'justified'] as const) {
    const expected = linesFor(exactLayout, name);
    const actual = linesFor(compact, name);
    expect(actual.map((line) => line.spans.map((span) => span.text))).toEqual(
      expected.map((line) => line.spans.map((span) => span.text))
    );
    for (const [lineIndex, line] of actual.entries()) {
      for (const [spanIndex, span] of line.spans.entries()) {
        // A spread of a compact box copies no coordinates, so this also proves the merge kept them.
        expect(isCompactBox(span.box)).toBe(true);
        const exact = expected[lineIndex]!.spans[spanIndex]!.box;
        for (const key of ['x', 'y', 'width', 'height'] as const) {
          expect(Math.abs(span.box[key] - exact[key])).toBeLessThan(0.001);
        }
      }
    }
  }
});
