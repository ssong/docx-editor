// One PDF export memory measurement, in a fresh Node process.
//
// `run.mjs` starts this file once per document and repetition with `--expose-gc`. It reads
// the built `@docx-editor.dev/docx-to-pdf` package, as a Node service does, and converts the
// document with `exportPdf`: open, layout, paint, and save in one call.
//
// `peakRssMb` after `export` is the number this probe is for: the process high-water mark
// of the whole conversion, including memory that each phase used and freed. The report also
// keeps the SHA-256 of the PDF, so two runs show whether a change altered the output bytes.
//
// Usage: node --expose-gc benchmarks/memory/pdf-probe.mjs <document.docx>

import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';

if (typeof globalThis.gc !== 'function') throw new Error('Run with node --expose-gc.');
const file = process.argv[2];
if (!file) throw new Error('Pass the document path.');

const mb = (bytes) => Math.round((bytes / 2 ** 20) * 100) / 100;
// Collect on a fresh macrotask, as `server-probe.mjs` does.
async function settle() {
  await new Promise((done) => setTimeout(done, 0));
  globalThis.gc();
  globalThis.gc();
  const usage = process.memoryUsage();
  return {
    heapUsedMb: mb(usage.heapUsed),
    externalMb: mb(usage.external),
    rssMb: mb(usage.rss),
    peakRssMb: mb(process.resourceUsage().maxRSS * 1024),
  };
}

const stages = [];
let base = await settle();
const { exportPdf } = await import('@docx-editor.dev/docx-to-pdf');
base = await settle();
stages.push({ stage: 'modules loaded', ms: 0, ...base, retainedMb: 0, deltaMb: 0 });
async function stage(name, fn) {
  const t0 = performance.now();
  const value = await fn();
  const ms = Math.round(performance.now() - t0);
  const memory = await settle();
  stages.push({
    stage: name,
    ms,
    ...memory,
    retainedMb: Math.round((memory.heapUsedMb - base.heapUsedMb) * 100) / 100,
    deltaMb: Math.round((memory.heapUsedMb - stages.at(-1).heapUsedMb) * 100) / 100,
  });
  return value;
}

const bytes = new Uint8Array(readFileSync(file));
// Only the packaged fonts, so the result does not depend on the fonts the machine has.
let result = await stage('export', () =>
  exportPdf(bytes, { useSystemFonts: false, fidelityPolicy: 'best-effort', timeoutMs: 3_600_000 })
);
const { pageCount, timings } = result;
const pdfBytes = result.bytes.byteLength;
const pdfSha256 = createHash('sha256').update(result.bytes).digest('hex');
await stage('release', () => {
  result = null;
});

console.log(
  JSON.stringify({
    file,
    bytes: bytes.length,
    pages: pageCount,
    pdfBytes,
    pdfSha256,
    node: process.version,
    ...timings,
    stages,
  })
);
