// Heap snapshots of the server-side stages, for finding what holds memory.
//
// Opens one document the same way `server-probe.mjs` does and writes a V8 heap snapshot
// after `open`, `layout`, and `release`, then for a fresh store before and after 200 edits. Load the snapshots in Chrome DevTools
// (Memory tab), or summarize them with `analyze-snapshot.mjs`.
//
// Usage: node --expose-gc benchmarks/memory/server-snapshot.mjs <document.docx> <out-dir>

import { mkdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { writeHeapSnapshot } from 'node:v8';

const [file, outDir] = process.argv.slice(2);
if (!file || !outDir) throw new Error('Pass the document path and an output directory.');
mkdirSync(outDir, { recursive: true });
const store = await import('@docx-editor.dev/core/store');
const exporter = await import('@docx-editor.dev/core/export');

const snapshot = async (name) => {
  // A fresh macrotask: see `settle` in server-probe.mjs.
  await new Promise((done) => setTimeout(done, 0));
  globalThis.gc();
  globalThis.gc();
  const path = writeHeapSnapshot(join(outDir, `${name}.heapsnapshot`));
  console.log(`${name}: ${(process.memoryUsage().heapUsed / 2 ** 20).toFixed(1)} MB → ${path}`);
};

await snapshot('0-baseline');
const bytes = new Uint8Array(readFileSync(file));
let view = (() => {
  const opened = store.openHeadlessDocument(bytes);
  if (!opened.ok) throw new Error(`open failed: ${opened.reason}`);
  return opened.view;
})();
await snapshot('1-open');
const result = exporter.openDocumentForExport(view);
if (!result.ok) throw new Error(`export open failed: ${result.reason}`);
let session = result.session;
let layout = await session.layout();
console.log(`${layout.pages.length} pages`);
await snapshot('2-layout');
session.dispose();
session = null;
layout = null;
view = null;
await snapshot('3-release');

// Edits: a fresh store, then single-character transactions, as in server-probe.mjs.
let tree = (() => {
  const loaded = store.readOoxmlPackage(bytes);
  if (!loaded.ok) throw new Error(`parse failed: ${loaded.reason}`);
  const main = loaded.package.parts.get(loaded.package.mainDocumentPart);
  return new store.TreePackageStore(loaded.package, store.normalizeParagraphIdentity(main));
})();
await snapshot('4-reopened');
const paragraphIds = [];
(function walk(node) {
  if (node.kind === 'paragraph') paragraphIds.push(node.id);
  for (const child of node.children ?? []) walk(child);
})(tree.bodyStore().part.root);
const edits = 200;
for (let i = 0; i < edits; i += 1) {
  const paragraphId = paragraphIds[Math.floor((i / edits) * paragraphIds.length)];
  tree
    .bodyStore()
    .transact((ctx) => ctx.apply({ op: 'insertText', paragraphId, offset: 0, text: 'x' }));
}
await snapshot(`5-edited-x${edits}`);
