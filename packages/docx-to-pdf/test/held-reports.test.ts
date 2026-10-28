/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/docx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { describe, expect, test } from 'bun:test';
import { PdfWorkLimitError, Work } from '../src/context.ts';

const work = () => new Work(new AbortController().signal);

describe('held reports', () => {
  test('reports made page by page come out phase by phase', () => {
    // Two phases over two pages. Held, they report as two passes over the whole document.
    const paged = work();
    paged.holdReports(2);
    for (const page of [0, 1]) {
      paged.reportPhase(0);
      paged.report('a', 'first phase', page);
      if (page === 1) paged.report('c', 'both', undefined, 'information');
      paged.reportPhase(1);
      paged.report('b', 'second phase', page);
      paged.report('c', 'both', undefined);
    }
    expect(paged.diagnostics).toEqual([]);
    paged.releaseReports();
    // `c` keeps the first report in phase order, the one from the first phase.
    expect(paged.diagnostics.map(({ code, pageIndex }) => [code, pageIndex])).toEqual([
      ['a', 0],
      ['a', 1],
      ['c', undefined],
      ['b', 0],
      ['b', 1],
    ]);
    expect(paged.diagnostics.find((d) => d.code === 'c')?.severity).toBe('information');

    // Reports after the release are not held.
    paged.report('d', 'after', 0);
    expect(paged.diagnostics.at(-1)?.code).toBe('d');
  });

  test('a report seen before holding is not repeated', () => {
    const paged = work();
    paged.report('a', 'early', 0);
    paged.holdReports(1);
    paged.report('a', 'early', 0);
    paged.releaseReports();
    expect(paged.diagnostics).toHaveLength(1);
  });

  test('held reports keep the diagnostic limit', () => {
    const paged = work();
    paged.holdReports(1);
    expect(() => {
      for (let index = 0; index <= 10_000; index += 1) paged.report('a', String(index));
    }).toThrow(PdfWorkLimitError);
  });
});
