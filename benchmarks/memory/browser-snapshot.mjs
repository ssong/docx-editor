// Heap snapshots of the browser stages, for finding what holds memory.
//
// Builds the benchmark page from the engine source without minification, so function and
// variable names in the snapshots match the code. Writes a snapshot after `open`, after `type`, and after
// `destroy`. Load them in Chrome DevTools (Memory tab), or summarize them with
// `analyze-snapshot.mjs`.
//
// Usage: node benchmarks/memory/browser-snapshot.mjs <document.docx> <out-dir>
//          [--edits 20] [--chromium <path>]

import { execFileSync } from 'node:child_process';
import { createWriteStream, mkdirSync, readFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { extname, join, resolve } from 'node:path';
import { parseArgs } from 'node:util';
import { chromium } from '@playwright/test';

const here = import.meta.dirname;
const root = resolve(here, '../..');
const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    edits: { type: 'string', default: '20' },
    chromium: { type: 'string', default: process.env.MEMORY_BENCH_CHROMIUM },
  },
});
const [file, outDir] = positionals;
if (!file || !outDir) throw new Error('Pass the document path and an output directory.');
mkdirSync(outDir, { recursive: true });

const pageDir = join(root, 'local', 'memory-bench', 'page-profile');
execFileSync(
  join(root, 'node_modules/.bin/vite'),
  ['build', '--config', join(here, 'browser/vite.config.mjs'), '--outDir', pageDir],
  { cwd: root, stdio: 'ignore', env: { ...process.env, MEMORY_BENCH_SOURCE: '1' } }
);
const types = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.wasm': 'application/wasm',
};
const server = createServer((req, res) => {
  const url = decodeURIComponent((req.url ?? '/').split('?')[0]);
  const path =
    url === '/document.docx' ? resolve(file) : join(pageDir, url === '/' ? 'index.html' : url);
  if (url !== '/document.docx' && !path.startsWith(pageDir)) return res.writeHead(403).end();
  try {
    const body = readFileSync(path);
    res.writeHead(200, { 'Content-Type': types[extname(path)] ?? 'application/octet-stream' });
    res.end(body);
  } catch {
    res.writeHead(404).end();
  }
});
await new Promise((done) => server.listen(0, '127.0.0.1', done));

const browser = await chromium.launch(values.chromium ? { executablePath: values.chromium } : {});
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  const cdp = await page.context().newCDPSession(page);
  const snapshot = async (name) => {
    await cdp.send('HeapProfiler.collectGarbage');
    await cdp.send('HeapProfiler.collectGarbage');
    const { usedSize } = await cdp.send('Runtime.getHeapUsage');
    const path = join(outDir, `${name}.heapsnapshot`);
    const stream = createWriteStream(path);
    const write = (message) => stream.write(message.chunk);
    cdp.on('HeapProfiler.addHeapSnapshotChunk', write);
    await cdp.send('HeapProfiler.takeHeapSnapshot', { reportProgress: false });
    cdp.off('HeapProfiler.addHeapSnapshotChunk', write);
    await new Promise((done) => stream.end(done));
    console.log(`${name}: ${(usedSize / 2 ** 20).toFixed(1)} MB → ${path}`);
  };
  await page.goto(`http://127.0.0.1:${server.address().port}/index.html`);
  await page.waitForFunction(() => !!window.__memoryBench);
  await page.evaluate(() => window.__memoryBench.loadFonts());
  await snapshot('0-fonts-loaded');
  const opened = await page.evaluate(() => window.__memoryBench.open('/document.docx'));
  console.log(`${opened.pages} pages`);
  await snapshot('1-open');
  await page.evaluate((n) => window.__memoryBench.type(n), Number(values.edits));
  await snapshot('2-typed');
  await page.evaluate(() => window.__memoryBench.destroy());
  await snapshot('3-destroyed');
} finally {
  await browser.close();
  server.close();
}
