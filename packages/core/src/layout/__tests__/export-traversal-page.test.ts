import { expect, test } from 'bun:test';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { forEachSemanticSpan, type SemanticSpanVisit } from '../export-traversal.ts';
import { noticeFixture } from './note-continuation-fixture.ts';

const walk = (visit: (visitor: (span: SemanticSpanVisit) => void) => void): SemanticSpanVisit[] => {
  const spans: SemanticSpanVisit[] = [];
  visit((span) => spans.push(span));
  return spans;
};

test('a walk page by page visits the same spans in the same order as one walk', () => {
  const { document, options } = noticeFixture('Continued');
  const layout = layoutSemanticDocument(document, 0, options);
  expect(layout.pages.length).toBeGreaterThan(1);
  for (const spanlessMarkers of [false, true]) {
    const whole = walk((visitor) => forEachSemanticSpan(layout, visitor, { spanlessMarkers }));
    const paged = layout.pages.flatMap((page) =>
      walk((visitor) => forEachSemanticSpan(layout, visitor, { spanlessMarkers, page }))
    );
    expect(paged).toEqual(whole);
    for (const page of layout.pages) {
      const own = walk((visitor) =>
        forEachSemanticSpan(layout, visitor, { spanlessMarkers, page })
      );
      expect(own.every((span) => span.page === page)).toBe(true);
      expect(own.length).toBeGreaterThan(0);
    }
  }
});
