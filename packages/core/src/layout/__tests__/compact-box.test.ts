import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { compactBox, isCompactBox, withBox } from '../compact-box.ts';
import { forEachSemanticSpan } from '../export-traversal.ts';
import { openDocumentForExport } from '../../export/export-session.ts';
import type { LayoutBox } from '../semantic-records.ts';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../../../../../e2e/fixtures/${name}`, import.meta.url)));

const fieldsOf = (box: LayoutBox): LayoutBox => ({
  x: box.x,
  y: box.y,
  width: box.width,
  height: box.height,
});

describe('compactBox', () => {
  test('rounds each coordinate to 1/20000 pt and reads back points', () => {
    const box = compactBox({ x: 72.123456, y: 0, width: 13.8, height: 1 / 3 });
    expect(isCompactBox(box)).toBe(true);
    expect(fieldsOf(box)).toEqual({ x: 72.12345, y: 0, width: 13.8, height: 0.33335 });
  });

  test('keeps whole twips exact', () => {
    for (let twips = -2000; twips <= 2000; twips += 7) {
      expect(compactBox({ x: twips / 20, y: 0, width: 0, height: 0 }).x).toBe(twips / 20);
    }
  });

  test('stores negative zero as zero', () => {
    expect(Object.is(compactBox({ x: -0.000001, y: -0, width: 0, height: 0 }).x, 0)).toBe(true);
  });

  test('keeps a box literal when a coordinate is off the integer grid', () => {
    const far = { x: 60_000, y: 0, width: 1, height: 1 };
    const infinite = { x: 0, y: 0, width: Number.POSITIVE_INFINITY, height: 1 };
    expect(compactBox(far)).toBe(far);
    expect(compactBox(infinite)).toBe(infinite);
  });

  test('returns a compact box unchanged', () => {
    const box = compactBox({ x: 1, y: 2, width: 3, height: 4 });
    expect(compactBox(box)).toBe(box);
  });

  test('serializes as a box literal', () => {
    const box = compactBox({ x: 1.5, y: 2, width: 3, height: 4 });
    expect(JSON.parse(JSON.stringify({ box }))).toEqual({
      box: { x: 1.5, y: 2, width: 3, height: 4 },
    });
  });

  test('withBox keeps the form of the box it copies', () => {
    const compact = compactBox({ x: 1, y: 2, width: 3, height: 4 });
    const moved = withBox(compact, { x: 5.5 });
    expect(isCompactBox(moved)).toBe(true);
    expect(fieldsOf(moved)).toEqual({ x: 5.5, y: 2, width: 3, height: 4 });

    const literal = { x: 1, y: 2, width: 3, height: 4 };
    const copied = withBox(literal, { width: 7 });
    expect(isCompactBox(copied)).toBe(false);
    expect(copied).toEqual({ x: 1, y: 2, width: 7, height: 4 });
  });
});

describe('compactLayoutBoxes', () => {
  async function spansOf(bytes: Uint8Array, compactLayoutBoxes: boolean) {
    const opened = openDocumentForExport(bytes, { compactLayoutBoxes });
    if (!opened.ok) throw new Error(`open failed: ${opened.reason}`);
    const layout = await opened.session.layout();
    const spans: { text: string; compact: boolean; box: LayoutBox }[] = [];
    forEachSemanticSpan(layout, (visit) => {
      spans.push({
        text: visit.span.text,
        compact: isCompactBox(visit.span.box),
        box: fieldsOf(visit.absoluteBox),
      });
    });
    opened.session.dispose();
    return spans;
  }

  for (const name of ['demo.docx', 'with-tables.docx', 'rtl-table-bidivisual.docx']) {
    test(`${name}: same spans within 0.001 pt, and the body spans are compact`, async () => {
      const bytes = fixture(name);
      const plain = await spansOf(bytes, false);
      const compact = await spansOf(bytes, true);
      expect(compact.length).toBe(plain.length);
      for (let index = 0; index < plain.length; index += 1) {
        const before = plain[index]!;
        const after = compact[index]!;
        expect(after.text).toBe(before.text);
        expect(before.compact).toBe(false);
        for (const key of ['x', 'y', 'width', 'height'] as const) {
          expect(Math.abs(after.box[key] - before.box[key])).toBeLessThan(0.001);
        }
      }
      const compactCount = compact.filter((span) => span.compact).length;
      expect(compactCount / compact.length).toBeGreaterThan(0.9);
    });
  }
});
