// Stable, record-only traversal for exporters and other non-DOM consumers.

import type {
  LineRecord,
  PageRecord,
  ParagraphFragmentRecord,
  SemanticLayout,
  SourceRange,
  StyleSpanRecord,
} from './semantic-records.ts';
import type { AnchoredDrawingRecord } from './drawing-layout.ts';
import { headerFooterAnchoredDrawingOrigin } from './header-footer-drawing-origin.ts';
import { everyStoryOrder } from './document-order.ts';
import { lineSegments } from './line-segments.ts';
import {
  forEachPageStory,
  forEachSemanticStory,
  forEachStoryParagraphFragment,
  type SemanticRootStoryKind,
  type SemanticStoryKind,
  type SemanticStoryVisit,
  type TextboxOwnerRecord,
} from './semantic-record-queries.ts';

export {
  forEachSemanticDrawing,
  forEachSemanticStory,
  forEachStoryParagraphFragment,
  type SemanticDrawingVisit,
  type SemanticDrawingLayer,
  type SemanticRootStoryKind,
  type SemanticStoryKind,
  type SemanticStoryVisit,
  type StoryDrawingContext,
  type StoryDrawingHost,
  type StoryParagraphFragmentContext,
  type TextboxOwnerRecord,
} from './semantic-record-queries.ts';

/** One span in the engine's published story order. @public */
export interface SemanticSpanVisit {
  readonly page: PageRecord;
  readonly story: SemanticStoryKind;
  /** Root story from which textbox descent began; equal to `story` outside textboxes. */
  readonly rootStory: SemanticRootStoryKind;
  /** Precise root host and absolute origin for story-relative geometry. */
  readonly root: SemanticStoryVisit;
  /** Absolute origin of the immediate root or textbox story containing this span. */
  readonly storyOrigin: Readonly<{ x: number; y: number }>;
  /** Absolute laid-out span bounds in page-stack coordinates. */
  readonly absoluteBox: import('./semantic-records.ts').LayoutBox;
  /** Owning note scope/area where applicable; null for body and page furniture. */
  readonly noteScopeId: string | null;
  readonly noteAreaKind: SemanticStoryVisit['noteAreaKind'];
  /** Zero outside a textbox, otherwise its bounded nesting depth. */
  readonly textboxDepth: number;
  /** Immediate textbox-owning drawing, anchored or inline, or null in the root story. */
  readonly textboxOwner: TextboxOwnerRecord | null;
  /** Root-to-leaf textbox owners, preserving drawing identity for future exporters. */
  readonly textboxPath: readonly TextboxOwnerRecord[];
  /** Enclosing published fragment; use paragraphId for the authored span owner. */
  readonly paragraph: ParagraphFragmentRecord;
  /** Authored paragraph owning this span, including spans merged into another fragment. */
  readonly paragraphId: string;
  readonly line: LineRecord;
  readonly span: StyleSpanRecord;
  /**
   * Model address for authored text. Projected atoms intentionally return null even though
   * their geometry record carries a range used internally by layout.
   */
  readonly sourceRange: SourceRange | null;
}

/** Return the model address exporters may use, excluding layout-projected atoms. @public */
export function exportSourceRangeOf(span: StyleSpanRecord): SourceRange | null {
  return span.projected === true ? null : span.range;
}

function visitStory(
  root: SemanticStoryVisit,
  paragraphOrder: ReadonlyMap<string, number>,
  visitor: (visit: SemanticSpanVisit) => void,
  spanlessMarkers = false
): void {
  const { page, story, host, noteScopeId, noteAreaKind } = root;
  const rootDrawingOrigin =
    story === 'header' || story === 'footer'
      ? (drawing: AnchoredDrawingRecord) =>
          headerFooterAnchoredDrawingOrigin(drawing, root.origin, {
            x: page.box.x,
            y: page.box.y,
          })
      : undefined;
  forEachStoryParagraphFragment(
    host,
    (block, textboxContext) => {
      const { textboxDepth, textboxOwner, textboxPath, storyOrigin } = textboxContext;
      const visitStoryKind = textboxDepth === 0 ? story : 'textbox';
      if (spanlessMarkers && block.marker && !block.lines.some((line) => line.spans.length > 0)) {
        // A marker with no text beside it: one empty span on the first line carries it.
        const first = block.lines[0];
        if (!first || !block.marker) return;
        const span: StyleSpanRecord = {
          range: {
            paragraphId: block.paragraphId,
            start: first.range.start,
            end: first.range.start,
          },
          text: '',
          props: [],
          style: block.marker.style,
          box: { ...first.box, width: 0 },
        };
        visitor({
          page,
          story: visitStoryKind,
          rootStory: story,
          root,
          storyOrigin,
          absoluteBox: Object.freeze({
            x: storyOrigin.x + span.box.x,
            y: storyOrigin.y + span.box.y,
            width: 0,
            height: span.box.height,
          }),
          noteScopeId,
          noteAreaKind,
          textboxDepth,
          textboxOwner,
          textboxPath,
          paragraph: block,
          paragraphId: block.paragraphId,
          line: first,
          span,
          sourceRange: null,
        });
        return;
      }
      for (const line of block.lines) {
        const segments = [...lineSegments(line)].sort(
          (left, right) =>
            (paragraphOrder.get(left.paragraphId) ?? Number.MAX_SAFE_INTEGER) -
            (paragraphOrder.get(right.paragraphId) ?? Number.MAX_SAFE_INTEGER)
        );
        for (const segment of segments) {
          for (const span of segment.spans) {
            visitor({
              page,
              story: visitStoryKind,
              rootStory: story,
              root,
              storyOrigin,
              absoluteBox: Object.freeze({
                x: storyOrigin.x + span.box.x,
                y: storyOrigin.y + span.box.y,
                width: span.box.width,
                height: span.box.height,
              }),
              noteScopeId,
              noteAreaKind,
              textboxDepth,
              textboxOwner,
              textboxPath,
              paragraph: block,
              paragraphId: segment.paragraphId,
              line,
              span,
              sourceRange: exportSourceRangeOf(span),
            });
          }
        }
      }
    },
    root.origin,
    rootDrawingOrigin
  );
}

/**
 * Visit every published span in page/story order without consulting the source package.
 * With `spanlessMarkers`, a paragraph fragment that publishes a list marker but no span, such
 * as an empty numbered paragraph, is visited once, in order, with an empty span on its first
 * line, so an exporter that paints markers beside spans paints its marker too.
 * With `page`, one of `layout.pages`, only that page's spans are visited, in the order the
 * whole-layout walk visits them, so an exporter can hold one page of visits at a time.
 * @public
 */
export function forEachSemanticSpan(
  layout: SemanticLayout,
  visitor: (visit: SemanticSpanVisit) => void,
  options: { readonly spanlessMarkers?: boolean; readonly page?: PageRecord } = {}
): void {
  const paragraphOrder = everyStoryOrderIndex(layout);
  const visit = (story: SemanticStoryVisit): void =>
    visitStory(story, paragraphOrder, visitor, options.spanlessMarkers === true);
  if (options.page) forEachPageStory(options.page, visit);
  else forEachSemanticStory(layout, visit);
}

const everyStoryOrderIndexCache = new WeakMap<SemanticLayout, ReadonlyMap<string, number>>();

/** Every story's paragraph order as a lookup, built once per layout for page-by-page walks. */
function everyStoryOrderIndex(layout: SemanticLayout): ReadonlyMap<string, number> {
  let index = everyStoryOrderIndexCache.get(layout);
  if (!index) {
    index = new Map(everyStoryOrder(layout).map((paragraphId, order) => [paragraphId, order]));
    everyStoryOrderIndexCache.set(layout, index);
  }
  return index;
}
