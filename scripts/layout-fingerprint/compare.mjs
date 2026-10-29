#!/usr/bin/env node
// Compares two layout fingerprint directories written by `run.mjs`: base, then head.
//
// Exit code 0 when every document matches, 1 when one differs, 2 on a usage error.
// See scripts/layout-fingerprint/README.md for what each mode accepts.

import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { parseArgs } from 'node:util';
import { gunzipSync } from 'node:zlib';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    tolerance: { type: 'string', default: '0' },
    spans: { type: 'string', default: 'exact' },
    'max-diffs': { type: 'string', default: '10' },
    markdown: { type: 'string' },
    help: { type: 'boolean', default: false },
  },
});

if (values.help || positionals.length !== 2) {
  console.log(`Usage: node scripts/layout-fingerprint/compare.mjs <base-dir> <head-dir> [options]

Options:
  --tolerance <pt>    allowed difference for every position and width, in points
                      (default: 0, which means exact)
  --spans <mode>      exact: same spans on every line (default)
                      merged: a head span may join adjacent base spans of one line, if
                      its text, start x and end x match the joined spans
  --max-diffs <n>     differences listed per document (default: 10)
  --markdown <file>   also write the report as Markdown (for a CI job summary)`);
  process.exit(values.help ? 0 : 2);
}

const tolerance = Number(values.tolerance);
if (!(tolerance >= 0)) throw new Error('--tolerance takes a number of points, 0 or more');
if (!['exact', 'merged'].includes(values.spans)) throw new Error('--spans takes exact or merged');
const maxDiffs = Number(values['max-diffs']);
const [baseDir, headDir] = positionals.map((p) => resolve(p));

const readManifest = (dir) => JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
const base = readManifest(baseDir);
const head = readManifest(headDir);
const readFingerprint = (dir, file) => JSON.parse(gunzipSync(readFileSync(join(dir, file))));

const near = (a, b) => (tolerance === 0 ? a === b : Math.abs(a - b) <= tolerance + 1e-9);
// A joined span's end is a sum of advances on one side and one x + width on the other, so
// exact mode still allows the rounding error of that sum.
const nearSum = (a, b) => Math.abs(a - b) <= Math.max(tolerance, 1e-6);
const pt = (n) => (Number.isInteger(n) ? String(n) : n.toFixed(4).replace(/0+$/, ''));
const quote = (text) => JSON.stringify(text.length > 60 ? `${text.slice(0, 57)}...` : text);
const lineText = (line) => line.spans.map((s) => s[0]).join('');
const describeLine = (line, page) =>
  `page ${page + 1}, ${line.s} line ${line.p} [${line.r[0]}, ${line.r[1]}] ${quote(lineText(line))}`;

/** Every line of a fingerprint in published order, with its page index. */
function flatLines(fingerprint) {
  const lines = [];
  fingerprint.pages.forEach((page, index) => {
    for (const line of page.lines) lines.push({ page: index, line });
  });
  return lines;
}

const SPAN_FIELDS = [
  [2, 'x'],
  [3, 'y'],
  [4, 'width'],
];

/** Differences between the spans of two lines that already match in story and range. */
function spanDiffs(baseSpans, headSpans) {
  if (values.spans === 'merged') return mergedSpanDiffs(baseSpans, headSpans);
  if (baseSpans.length !== headSpans.length) {
    return [`${baseSpans.length} spans → ${headSpans.length} spans`];
  }
  const diffs = [];
  baseSpans.forEach((b, i) => {
    const h = headSpans[i];
    if (b[0] !== h[0] || b[1] !== h[1] || b[5] !== h[5]) {
      diffs.push(`span ${i + 1} ${quote(b[0])} at ${b[1]} → ${quote(h[0])} at ${h[1]}`);
      return;
    }
    for (const [index, label] of SPAN_FIELDS) {
      if (!near(b[index], h[index])) {
        diffs.push(`span ${i + 1} ${quote(b[0])} ${label} ${pt(b[index])} → ${pt(h[index])}`);
      }
    }
  });
  return diffs;
}

/**
 * A head span may cover several adjacent base spans. It must start where the first starts
 * (offset, x, y) and end where the last ends (x + width), and carry the joined text.
 */
function mergedSpanDiffs(baseSpans, headSpans) {
  const diffs = [];
  let next = 0;
  for (const [i, h] of headSpans.entries()) {
    const start = next;
    let text = '';
    while (next < baseSpans.length && text.length < h[0].length) text += baseSpans[next++][0];
    const joined = baseSpans.slice(start, next);
    const [first, last] = [joined[0], joined.at(-1)];
    if (!first || text !== h[0] || first[1] !== h[1] || joined.some((s) => s[5] !== h[5])) {
      diffs.push(`span ${i + 1} ${quote(h[0])} at ${h[1]} does not join base spans`);
      return diffs;
    }
    if (!near(first[2], h[2]))
      diffs.push(`span ${i + 1} ${quote(h[0])} x ${pt(first[2])} → ${pt(h[2])}`);
    if (!near(first[3], h[3]))
      diffs.push(`span ${i + 1} ${quote(h[0])} y ${pt(first[3])} → ${pt(h[3])}`);
    const [baseEnd, headEnd] = [last[2] + last[4], h[2] + h[4]];
    if (!nearSum(baseEnd, headEnd)) {
      diffs.push(`span ${i + 1} ${quote(h[0])} end x ${pt(baseEnd)} → ${pt(headEnd)}`);
    }
  }
  if (next < baseSpans.length)
    diffs.push(`base span ${quote(baseSpans[next][0])} has no head span`);
  return diffs;
}

const LINE_FIELDS = [
  ['x', 'x'],
  ['y', 'baseline y'],
  ['w', 'width'],
  ['h', 'height'],
];

/** Differences for one document, at most `maxDiffs` of them, plus a short summary. */
function compareDocument(b, h) {
  if (b.error || h.error) {
    if (b.error === h.error) return { summary: null, diffs: [] };
    return {
      summary: 'layout result changed',
      diffs: [`${b.error ?? 'laid out'} → ${h.error ?? 'laid out'}`],
    };
  }
  const diffs = [];
  const push = (text) => diffs.length < maxDiffs && diffs.push(text);
  let summary = b.pageCount !== h.pageCount ? `${b.pageCount} → ${h.pageCount} pages` : null;
  if (b.sha256 !== h.sha256) push('the two sides laid out different bytes');
  for (let i = 0; i < Math.min(b.pages.length, h.pages.length); i += 1) {
    const [bp, hp] = [b.pages[i], h.pages[i]];
    if (bp.w !== hp.w || bp.h !== hp.h) {
      push(`page ${i + 1} size ${bp.w} x ${bp.h} → ${hp.w} x ${hp.h}`);
    }
  }
  const baseLines = flatLines(b);
  const headLines = flatLines(h);
  let differing = 0;
  for (let i = 0; i < Math.max(baseLines.length, headLines.length); i += 1) {
    const bl = baseLines[i];
    const hl = headLines[i];
    if (!bl || !hl) {
      push(
        `${bl ? 'base' : 'head'} has ${Math.abs(baseLines.length - headLines.length)} more lines, from ${describeLine((bl ?? hl).line, (bl ?? hl).page)}`
      );
      differing += 1;
      break;
    }
    const [x, y] = [bl.line, hl.line];
    if (x.s !== y.s || x.p !== y.p || x.r[0] !== y.r[0] || x.r[1] !== y.r[1]) {
      // A different line break shifts every later line, so the rest would only repeat it.
      push(
        `line ${i + 1} breaks differently: ${describeLine(x, bl.page)} → ${describeLine(y, hl.page)}`
      );
      summary ??= 'line breaks differ';
      differing += 1;
      break;
    }
    const found = [];
    if (bl.page !== hl.page) found.push(`page ${bl.page + 1} → ${hl.page + 1}`);
    for (const [key, label] of LINE_FIELDS) {
      if (!near(x[key], y[key])) found.push(`${label} ${pt(x[key])} → ${pt(y[key])}`);
    }
    if (values.spans === 'merged' && lineText(x) !== lineText(y)) {
      found.push(`text ${quote(lineText(x))} → ${quote(lineText(y))}`);
    } else {
      found.push(...spanDiffs(x.spans, y.spans));
    }
    if (!drawingsNear(x.drawings, y.drawings)) {
      found.push('inline drawings moved');
    }
    if (found.length > 0) {
      differing += 1;
      push(
        `${describeLine(x, bl.page)}: ${found.slice(0, 4).join('; ')}${found.length > 4 ? `; ${found.length - 4} more` : ''}`
      );
    }
  }
  if (b.orphanSpans !== h.orphanSpans)
    push(`spans outside any line ${b.orphanSpans ?? 0} → ${h.orphanSpans ?? 0}`);
  if (diffs.length === 0) return { summary: null, diffs };
  summary ??= `${differing} line${differing === 1 ? '' : 's'} differ${differing === 1 ? 's' : ''}`;
  return { summary, diffs };
}

function drawingsNear(a = [], b = []) {
  return (
    a.length === b.length &&
    a.every((d, i) => d.every((v, j) => (j === 0 ? v === b[i][j] : near(v, b[i][j]))))
  );
}

const key = (doc) => `${doc.measurer}\u0000${doc.name}`;
const headByKey = new Map(head.documents.map((doc) => [key(doc), doc]));
const baseByKey = new Map(base.documents.map((doc) => [key(doc), doc]));
const results = [];
for (const doc of head.documents) {
  const other = baseByKey.get(key(doc));
  if (!other) {
    results.push({ doc, summary: 'missing on base', diffs: [] });
    continue;
  }
  const { summary, diffs } = compareDocument(
    readFingerprint(baseDir, other.file),
    readFingerprint(headDir, doc.file)
  );
  results.push({ doc, summary, diffs });
}
for (const doc of base.documents) {
  if (!headByKey.has(key(doc))) results.push({ doc, summary: 'missing on head', diffs: [] });
}

const changed = results.filter((r) => r.summary);
const mode = `${tolerance === 0 ? 'exact positions' : `positions within ${tolerance} pt`}, ${values.spans} spans`;
const title = changed.length
  ? `Layout differs in ${changed.length} of ${results.length} documents (${mode})`
  : `Layout matches in all ${results.length} documents (${mode})`;

const text = [title, ''];
for (const { doc, summary, diffs } of changed) {
  text.push(`✗ ${doc.measurer} ${doc.name}: ${summary}`);
  for (const diff of diffs) text.push(`    ${diff}`);
}
console.log(text.join('\n'));

if (values.markdown) {
  const md = [`### ${changed.length ? '✗' : '✓'} ${title}`, ''];
  md.push(
    `Base \`${base.commit?.slice(0, 7) ?? base.root}\`, head \`${head.commit?.slice(0, 7) ?? head.root}\`.`,
    ''
  );
  if (changed.length) {
    md.push('| Measurer | Document | Change |', '| --- | --- | --- |');
    for (const { doc, summary } of changed)
      md.push(`| ${doc.measurer} | \`${doc.name}\` | ${summary} |`);
    md.push('', '<details><summary>First differences per document</summary>', '', '```');
    for (const { doc, diffs } of changed) {
      md.push(`${doc.measurer} ${doc.name}`, ...diffs.map((d) => `    ${d}`));
    }
    md.push('```', '', '</details>');
  }
  writeFileSync(values.markdown, `${md.join('\n')}\n`);
}

process.exit(changed.length ? 1 : 0);
