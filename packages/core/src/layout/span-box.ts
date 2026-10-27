// Where a published span sits in its story.

import type { LayoutBox } from './semantic-records.ts';

/**
 * A span's box in its story's coordinates.
 *
 * A span box is horizontally in story coordinates and vertically relative to its line: `y` is
 * the offset from the line top, zero for ordinary text. That lets a laid-out line share its
 * spans with the paragraph cache, and lets a line move down a page without copying them. Add
 * the line top to get the span's place in the story.
 * @public
 */
export function spanStoryBox(
  line: { readonly box: Pick<LayoutBox, 'y'> },
  span: { readonly box: LayoutBox }
): LayoutBox {
  return {
    x: span.box.x,
    y: line.box.y + span.box.y,
    width: span.box.width,
    height: span.box.height,
  };
}
