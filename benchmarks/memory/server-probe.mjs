// One server-side memory measurement, in a fresh Node process.
//
// `run.mjs` starts this file once per document and repetition with `--expose-gc`, so no
// stage inherits heap, JIT state, or module caches from another document. It reads the
// built `@docx-editor.dev/core` package, the code a Node service installs.
//
// Each stage keeps what a real caller keeps, forces two full collections, and records the
// live heap. `retained` is what the document costs while it stays open. `peakRssMb` is the
// process high-water mark, which includes the garbage each stage created and freed.
//
// Usage: node --expose-gc benchmarks/memory/server-probe.mjs <document.docx> [--edits 200]
//        [--layout-spans word|run]

import { readFileSync } from 'node:fs';
import { parseArgs } from 'node:util';

if (typeof globalThis.gc !== 'function') throw new Error('Run with node --expose-gc.');

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    edits: { type: 'string', default: '200' },
    'layout-spans': { type: 'string', default: 'word' },
  },
});
const file = positionals[0];
if (!file) throw new Error('Pass the document path.');
const editCount = Number(values.edits);

const mb = (bytes) => Math.round((bytes / 2 ** 20) * 100) / 100;
// Collect on a fresh macrotask. Inside the job that resumed an `await`, V8 keeps the
// awaited value and every WeakRef target touched in the job alive, so a reading taken there
// counts memory that is already garbage.
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
let base = null;
async function stage(name, fn) {
  const t0 = performance.now();
  const value = await fn();
  const ms = Math.round(performance.now() - t0);
  const memory = await settle();
  stages.push({
    stage: name,
    ms,
    ...memory,
    retainedMb: Math.round((memory.heapUsedMb - (base?.heapUsedMb ?? 0)) * 100) / 100,
    deltaMb: Math.round((memory.heapUsedMb - stages.at(-1).heapUsedMb) * 100) / 100,
  });
  return value;
}

base = await settle();
const store = await import('@docx-editor.dev/core/store');
const exporter = await import('@docx-editor.dev/core/export');
base = await settle();
stages.push({ stage: 'modules loaded', ms: 0, ...base, retainedMb: 0, deltaMb: 0 });

const bytes = new Uint8Array(readFileSync(file));

// 1. Open: bytes → canonical tree → store. What a server holds for an open document.
let opened = await stage('open', () => {
  const result = store.openHeadlessDocument(bytes);
  if (!result.ok) throw new Error(`open failed: ${result.reason}`);
  return result.view;
});

// 2. Layout: the export session over the open view (the PDF and page-count path).
let session = null;
let pages = 0;
await stage('layout', async () => {
  const result = exporter.openDocumentForExport(opened, {
    reuseAcrossRevisions: false,
    spans: values['layout-spans'],
  });
  if (!result.ok) throw new Error(`export open failed: ${result.reason}`);
  session = result.session;
  const layout = await session.layout();
  pages = layout.pages.length;
});

// 3. Release: dispose the session and drop the view. Should fall back to the module baseline.
await stage('release', () => {
  session.dispose();
  session = null;
  opened = null;
});

// 4. Edit: a fresh store, then single-character transactions. Each one is an undo step, so
//    this shows what the undo history keeps per edit.
let tree = null;
await stage('reopen store', () => {
  const loaded = store.readOoxmlPackage(bytes);
  if (!loaded.ok) throw new Error(`parse failed: ${loaded.reason}`);
  const main = loaded.package.parts.get(loaded.package.mainDocumentPart);
  tree = new store.TreePackageStore(loaded.package, store.normalizeParagraphIdentity(main));
});
const paragraphIds = [];
(function walk(node) {
  if (node.kind === 'paragraph') paragraphIds.push(node.id);
  for (const child of node.children ?? []) walk(child);
})(tree.bodyStore().part.root);
await stage(`edit x${editCount}`, () => {
  const body = tree.bodyStore();
  for (let i = 0; i < editCount; i += 1) {
    const paragraphId = paragraphIds[Math.floor((i / editCount) * paragraphIds.length)];
    const ok = body.transact((ctx) =>
      ctx.apply({ op: 'insertText', paragraphId, offset: 0, text: 'x' })
    );
    if (ok?.ok === false) throw new Error(`edit refused: ${ok.reason}`);
  }
});

// 5. Save: serialize the edited package. Peak RSS shows the transient cost of the zip.
let saved = null;
await stage('save', () => {
  saved = store.writeOoxmlPackage(tree.currentPackage());
});

console.log(
  JSON.stringify({
    file,
    bytes: bytes.length,
    savedBytes: saved.length,
    pages,
    paragraphs: paragraphIds.length,
    node: process.version,
    stages,
  })
);
