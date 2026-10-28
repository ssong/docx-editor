// Span boxes that store their coordinates as small integers.
//
// A box literal `{ x, y, width, height }` holds four fractional numbers, and V8 keeps each one
// as a separate heap number: about 114 bytes per box on a server and 76 in a browser. A span
// box is published once per word, so these boxes were a tenth of a laid-out document. Rounded
// to a fixed grid, the same coordinates are integers that V8 stores inside the object.
//
// Only published spans take this form, and only when the paragraph cache asks for it
// (`compactBoxes`). Layout breaks a paragraph with the exact values; the cache then rounds the
// closed lines once, before placement reads them.

import type { LayoutBox } from './semantic-records.ts';

/**
 * Grid units per point: 1/20000 pt, or 1/1000 twip.
 *
 * Fine enough that rounding moves a coordinate by at most 0.000025 pt, so no threshold layout
 * compares a placed span against changes its answer. A whole number of twips stays exact.
 */
const UNITS_PER_POINT = 20_000;

/**
 * The largest integer V8 stores inside an object with pointer compression (a Smi).
 *
 * A coordinate past it, about 53,000 pt, would come back as a heap number and cost more than
 * the literal; such a box stays a literal. Page-local coordinates never come near it.
 */
const MAX_UNITS = 2 ** 30 - 1;

/** A coordinate in grid units, or NaN when it is not finite or not small enough. */
function units(value: number): number {
  // `+ 0` turns -0 into 0: V8 stores -0 as a heap number.
  const scaled = Math.round(value * UNITS_PER_POINT) + 0;
  return scaled >= -MAX_UNITS && scaled <= MAX_UNITS ? scaled : Number.NaN;
}

/**
 * A layout box that keeps its coordinates as integer grid units.
 *
 * The coordinates read as points through prototype getters, so `box.x` works as it does on a
 * literal. Object spread copies only own properties, so `{ ...box }` loses them: derive a box
 * with {@link withBox} instead.
 */
class CompactLayoutBox implements LayoutBox {
  readonly #x: number;
  readonly #y: number;
  readonly #width: number;
  readonly #height: number;

  constructor(x: number, y: number, width: number, height: number) {
    this.#x = x;
    this.#y = y;
    this.#width = width;
    this.#height = height;
  }

  static isCompact(box: LayoutBox): box is CompactLayoutBox {
    return #x in box;
  }

  get x(): number {
    return this.#x / UNITS_PER_POINT;
  }

  get y(): number {
    return this.#y / UNITS_PER_POINT;
  }

  get width(): number {
    return this.#width / UNITS_PER_POINT;
  }

  get height(): number {
    return this.#height / UNITS_PER_POINT;
  }

  toJSON(): LayoutBox {
    return { x: this.x, y: this.y, width: this.width, height: this.height };
  }
}

// Published layouts are immutable: no caller may replace a getter that every compact box reads.
Object.freeze(CompactLayoutBox.prototype);
Object.freeze(CompactLayoutBox);

/** Whether a box keeps its coordinates as integer grid units. */
export function isCompactBox(box: object): boolean {
  return CompactLayoutBox.isCompact(box as LayoutBox);
}

/**
 * The box with its coordinates rounded to the compact grid, or the box itself when it is
 * already compact or a coordinate is too large to store as a small integer.
 */
export function compactBox(box: LayoutBox): LayoutBox {
  if (CompactLayoutBox.isCompact(box)) return box;
  const x = units(box.x);
  const y = units(box.y);
  const width = units(box.width);
  const height = units(box.height);
  // NaN marks a coordinate outside the grid. Any NaN makes the sum NaN.
  if (Number.isNaN(x + y + width + height)) return box;
  return new CompactLayoutBox(x, y, width, height);
}

/**
 * A copy of `box` with some coordinates replaced, in the same form: compact stays compact.
 *
 * Use it instead of `{ ...box, x }` on a span box, because a spread copies nothing from a
 * compact box. A literal box is spread as before, so its other properties carry over.
 */
export function withBox(box: LayoutBox, patch: Partial<LayoutBox>): LayoutBox {
  if (!CompactLayoutBox.isCompact(box)) return { ...box, ...patch };
  return compactBox({
    x: patch.x ?? box.x,
    y: patch.y ?? box.y,
    width: patch.width ?? box.width,
    height: patch.height ?? box.height,
  });
}
