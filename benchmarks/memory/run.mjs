#!/usr/bin/env node
// Memory benchmark for `@docx-editor.dev/core`: server-side (Node) and in-browser (Chromium).
//
// Opens a range of document sizes and records how much memory each stage keeps alive:
// open, layout, scroll, edit, save, and release. Results go to `benchmarks/memory/results/`
// so a later run can compare against them. See benchmarks/memory/README.md.

import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { createServer } from 'node:http';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { cpus, totalmem, platform, arch } from 'node:os';
import { basename, extname, join, resolve } from 'node:path';
import { parseArgs } from 'node:util';
import { syntheticDocx } from './fixtures.mjs';

const here = import.meta.dirname;
const root = resolve(here, '../..');
const work = join(root, 'local', 'memory-bench');

const { values } = parseArgs({
  options: {
    sizes: { type: 'string', default: '10,50,100,250,500,1000' },
    real: { type: 'string', default: 'e2e/fixtures/demo.docx,e2e/fixtures/typing-perf-521pp.docx' },
    runs: { type: 'string', default: '3' },
    target: { type: 'string', default: 'all' },
    'server-edits': { type: 'string', default: '200' },
    'browser-edits': { type: 'string', default: '50' },
    fonts: { type: 'string', default: 'shaped' },
    'layout-spans': { type: 'string', default: 'word' },
    out: { type: 'string' },
    compare: { type: 'string' },
    'no-build': { type: 'boolean', default: false },
    chromium: { type: 'string', default: process.env.MEMORY_BENCH_CHROMIUM },
    help: { type: 'boolean', default: false },
  },
});

if (values.help) {
  console.log(`Usage: bun run bench:memory [options]

Options:
  --sizes <a,b,...>     synthetic document sizes, in content units (about one page each)
                        (default: 10,50,100,250,500,1000)
  --real <paths>        comma-separated real documents to add ('' for none)
                        (default: e2e/fixtures/demo.docx,e2e/fixtures/typing-perf-521pp.docx)
  --runs <n>            repetitions per document; the report keeps the median (default: 3)
  --target <which>      all | server | browser | pdf (default: all; all includes pdf
                        when @docx-editor.dev/docx-to-pdf is built)
  --server-edits <n>    single-character edits in the server edit stage (default: 200)
  --browser-edits <n>   typed characters in the browser edit stage (default: 50)
  --fonts <mode>        shaped (packaged default fonts, HarfBuzz) | fixed (default: shaped)
  --layout-spans <m>    the layout option for spans per line: word | run (default: word)
  --out <file>          write the JSON report here
                        (default: benchmarks/memory/results/<date>-<commit>.json)
  --compare <file>      print the change against an earlier report
  --no-build            reuse the built benchmark page
  --chromium <path>     Chromium executable (default: Playwright's own, or
                        the MEMORY_BENCH_CHROMIUM environment variable)

Build the packages first: bun run --filter '@docx-editor.dev/i18n' build &&
bun run --filter '@docx-editor.dev/core' build && bun run --filter '@docx-editor.dev/fonts' build.
The pdf target also needs bun run build:pdf.`);
  process.exit(0);
}

const runs = Number(values.runs);
const pdfBuilt = existsSync(join(root, 'packages/docx-to-pdf/dist/index.js'));
const targets =
  values.target === 'all' ? ['server', 'browser', ...(pdfBuilt ? ['pdf'] : [])] : [values.target];
if (!targets.every((t) => ['server', 'browser', 'pdf'].includes(t))) {
  throw new Error('--target takes all, server, browser, or pdf');
}
if (targets.includes('pdf') && !pdfBuilt) {
  throw new Error('Build @docx-editor.dev/docx-to-pdf first: bun run build:pdf.');
}
if (!['shaped', 'fixed'].includes(values.fonts)) throw new Error('--fonts takes shaped or fixed');
if (!['word', 'run'].includes(values['layout-spans']))
  throw new Error('--layout-spans takes word or run');
if (!existsSync(join(root, 'packages/core/dist/store.js'))) {
  throw new Error('Build @docx-editor.dev/core first (see --help).');
}

// ── Documents ──
const fixtureDir = join(work, 'fixtures');
mkdirSync(fixtureDir, { recursive: true });
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const documents = [];
for (const units of values.sizes.split(',').filter(Boolean).map(Number)) {
  const bytes = syntheticDocx(units);
  const path = join(fixtureDir, `synthetic-${units}.docx`);
  writeFileSync(path, bytes);
  documents.push({ name: `synthetic-${units}`, path, bytes: bytes.length, sha256: sha256(bytes) });
}
for (const real of values.real.split(',').filter(Boolean)) {
  const path = resolve(root, real);
  const bytes = readFileSync(path);
  documents.push({
    name: basename(real, extname(real)),
    path,
    bytes: bytes.length,
    sha256: sha256(bytes),
  });
}

const median = (list) => {
  const sorted = [...list].sort((a, b) => a - b);
  const mid = sorted.length >> 1;
  return sorted.length % 2
    ? sorted[mid]
    : Math.round(((sorted[mid - 1] + sorted[mid]) / 2) * 100) / 100;
};
/** Merge repeated runs stage by stage, keeping the median of each numeric field. */
function mergeRuns(samples) {
  const first = samples[0];
  const merged = { ...first };
  for (const key of Object.keys(first)) {
    if (typeof first[key] === 'number') merged[key] = median(samples.map((s) => s[key]));
  }
  merged.stages = first.stages.map((stage, index) => {
    const out = { ...stage };
    for (const key of Object.keys(stage)) {
      if (typeof stage[key] === 'number')
        out[key] = median(samples.map((s) => s.stages[index][key]));
    }
    return out;
  });
  return merged;
}

// ── Server ──
function probe(script, args) {
  return new Promise((done, fail) => {
    const child = spawn(process.execPath, ['--expose-gc', join(here, script), ...args], {
      cwd: root,
      stdio: ['ignore', 'pipe', 'inherit'],
    });
    let out = '';
    child.stdout.on('data', (chunk) => (out += chunk));
    child.on('exit', (code) =>
      code === 0
        ? done(JSON.parse(out.trim().split('\n').at(-1)))
        : fail(new Error(`probe exited ${code}`))
    );
  });
}

async function runServer() {
  const results = [];
  for (const doc of documents) {
    const samples = [];
    for (let i = 0; i < runs; i += 1)
      samples.push(
        await probe('server-probe.mjs', [
          doc.path,
          '--edits',
          values['server-edits'],
          '--layout-spans',
          values['layout-spans'],
        ])
      );
    const merged = mergeRuns(samples);
    delete merged.file;
    results.push({ document: doc.name, ...merged });
    const open = merged.stages.find((s) => s.stage === 'open');
    const layout = merged.stages.find((s) => s.stage === 'layout');
    console.log(
      `server  ${doc.name.padEnd(24)} ${String(merged.pages).padStart(5)} pages  open ${open.retainedMb} MB  layout ${layout.retainedMb} MB  peak RSS ${Math.max(...merged.stages.map((s) => s.peakRssMb))} MB`
    );
  }
  return results;
}

// ── PDF export ──
async function runPdf() {
  const results = [];
  for (const doc of documents) {
    const samples = [];
    for (let i = 0; i < runs; i += 1) samples.push(await probe('pdf-probe.mjs', [doc.path]));
    if (new Set(samples.map((s) => s.pdfSha256)).size > 1) {
      throw new Error(`${doc.name}: the PDF bytes differ between runs`);
    }
    const merged = mergeRuns(samples);
    delete merged.file;
    results.push({ document: doc.name, ...merged });
    const exported = merged.stages.find((s) => s.stage === 'export');
    console.log(
      `pdf     ${doc.name.padEnd(24)} ${String(merged.pages).padStart(5)} pages  export ${exported.ms} ms  peak RSS ${exported.peakRssMb} MB  PDF ${merged.pdfSha256.slice(0, 12)}`
    );
  }
  return results;
}

// ── Browser ──
function serveDirectory(dir, docs) {
  const types = {
    '.html': 'text/html',
    '.js': 'text/javascript',
    '.css': 'text/css',
    '.wasm': 'application/wasm',
  };
  const byName = new Map(docs.map((d) => [`/docs/${d.name}.docx`, d.path]));
  const server = createServer((req, res) => {
    const url = decodeURIComponent((req.url ?? '/').split('?')[0]);
    const file = byName.get(url) ?? join(dir, url === '/' ? 'index.html' : url);
    if (!file.startsWith(dir) && !byName.has(url)) return res.writeHead(403).end();
    try {
      const body = readFileSync(file);
      res.writeHead(200, {
        'Content-Type': types[extname(file)] ?? 'application/octet-stream',
        // Cross-origin isolation enables performance.measureUserAgentSpecificMemory(), the
        // one reading that includes DOM and layout objects, not only the JavaScript heap.
        'Cross-Origin-Opener-Policy': 'same-origin',
        'Cross-Origin-Embedder-Policy': 'require-corp',
      });
      res.end(body);
    } catch {
      res.writeHead(404).end();
    }
  });
  return new Promise((done) => server.listen(0, '127.0.0.1', () => done(server)));
}

const launch = (chromium) =>
  chromium.launch({
    args: ['--enable-precise-memory-info', '--enable-blink-features=ForceEagerMeasureMemory'],
    ...(values.chromium ? { executablePath: values.chromium } : {}),
  });

async function browserRun(chromium, origin, doc) {
  const browser = await launch(chromium);
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
    const cdp = await page.context().newCDPSession(page);
    await cdp.send('Performance.enable');
    const stages = [];
    let base = null;
    const measure = async (stage, ms, extra = {}) => {
      await cdp.send('HeapProfiler.collectGarbage');
      await cdp.send('HeapProfiler.collectGarbage');
      const { usedSize } = await cdp.send('Runtime.getHeapUsage');
      const { metrics } = await cdp.send('Performance.getMetrics');
      const metric = (name) => metrics.find((m) => m.name === name)?.value ?? 0;
      const heapUsedMb = Math.round((usedSize / 2 ** 20) * 100) / 100;
      const pageMemory = await page.evaluate(async () => {
        if (!self.crossOriginIsolated) return null;
        const result = await performance.measureUserAgentSpecificMemory();
        const byType = {};
        for (const entry of result.breakdown) {
          for (const type of entry.types) byType[type] = (byType[type] ?? 0) + entry.bytes;
        }
        return { total: result.bytes, dom: byType.DOM ?? 0 };
      });
      const mbOf = (bytes) => Math.round((bytes / 2 ** 20) * 100) / 100;
      base ??= heapUsedMb;
      stages.push({
        stage,
        ms,
        heapUsedMb,
        retainedMb: Math.round((heapUsedMb - base) * 100) / 100,
        deltaMb: Math.round((heapUsedMb - (stages.at(-1)?.heapUsedMb ?? heapUsedMb)) * 100) / 100,
        heapTotalMb: mbOf(metric('JSHeapTotalSize')),
        ...(pageMemory ? { pageTotalMb: mbOf(pageMemory.total), domMb: mbOf(pageMemory.dom) } : {}),
        domNodes: metric('Nodes'),
        listeners: metric('JSEventListeners'),
        ...extra,
      });
    };
    await page.goto(`${origin}/index.html`);
    await page.waitForFunction(() => !!window.__memoryBench);
    await measure('page loaded', 0);
    if (values.fonts === 'shaped') {
      await page.evaluate(() => window.__memoryBench.loadFonts());
      await measure('fonts loaded', 0);
      // The fonts are the host's, loaded once for every document: count from here.
      base = stages.at(-1).heapUsedMb;
    }
    const opened = await page.evaluate(
      ([url, spans]) => window.__memoryBench.open(url, spans),
      [`/docs/${doc.name}.docx`, values['layout-spans']]
    );
    await measure('open', opened.ms);
    const scrolled = await page.evaluate(() => window.__memoryBench.scrollThrough());
    await measure('scroll through', scrolled.ms);
    const edits = Number(values['browser-edits']);
    const typed = await page.evaluate((n) => window.__memoryBench.type(n), edits);
    await measure(`type x${edits}`, typed.ms);
    await page.evaluate(() => window.__memoryBench.destroy());
    await measure('destroy', 0);
    return { pages: opened.pages, stages };
  } finally {
    await browser.close();
  }
}

async function runBrowser() {
  const { chromium } = await import('@playwright/test');
  const pageDir = join(work, 'page');
  if (!values['no-build'] || !existsSync(join(pageDir, 'index.html'))) {
    execFileSync(
      join(root, 'node_modules/.bin/vite'),
      ['build', '--config', join(here, 'browser/vite.config.mjs'), '--outDir', pageDir],
      { cwd: root, stdio: 'ignore' }
    );
  }
  const server = await serveDirectory(pageDir, documents);
  const origin = `http://127.0.0.1:${server.address().port}`;
  const results = [];
  try {
    for (const doc of documents) {
      const samples = [];
      for (let i = 0; i < runs; i += 1) samples.push(await browserRun(chromium, origin, doc));
      const merged = mergeRuns(samples);
      results.push({ document: doc.name, ...merged });
      const open = merged.stages.find((s) => s.stage === 'open');
      const scroll = merged.stages.find((s) => s.stage === 'scroll through');
      console.log(
        `browser ${doc.name.padEnd(24)} ${String(merged.pages).padStart(5)} pages  open ${open.retainedMb} MB  after scroll ${scroll.retainedMb} MB  DOM ${scroll.domNodes}`
      );
    }
  } finally {
    server.close();
  }
  return results;
}

// ── Report ──
const git = (...args) => {
  try {
    return execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
  } catch {
    return null;
  }
};
const commit = git('rev-parse', '--short', 'HEAD');
const report = {
  schema: 1,
  date: new Date().toISOString(),
  commit,
  dirty: !!git('status', '--porcelain', '--', 'packages/core', 'packages/docx-to-pdf'),
  environment: {
    platform: `${platform()} ${arch()}`,
    cpu: cpus()[0]?.model ?? 'unknown',
    cpus: cpus().length,
    memoryGb: Math.round(totalmem() / 2 ** 30),
    node: process.version,
  },
  options: {
    runs,
    fonts: values.fonts,
    layoutSpans: values['layout-spans'],
    serverEdits: Number(values['server-edits']),
    browserEdits: Number(values['browser-edits']),
  },
  documents: documents.map(({ path, ...doc }) => doc),
};
if (targets.includes('server')) report.server = await runServer();
if (targets.includes('pdf')) report.pdf = await runPdf();
if (targets.includes('browser')) {
  report.browser = await runBrowser();
  const { chromium } = await import('@playwright/test');
  const browser = await launch(chromium);
  report.environment.chromium = browser.version();
  await browser.close();
}

const out =
  values.out ?? join(here, 'results', `${report.date.slice(0, 10)}-${commit ?? 'nocommit'}.json`);
mkdirSync(resolve(out, '..'), { recursive: true });
writeFileSync(out, `${JSON.stringify(report, null, 2)}\n`);
console.log(`\nwrote ${out}`);

if (values.compare) {
  const before = JSON.parse(readFileSync(values.compare, 'utf8'));
  console.log(`\nChange against ${values.compare} (${before.commit}):`);
  for (const target of ['server', 'browser', 'pdf']) {
    // A PDF export keeps little once it returns; its cost is the peak while it runs.
    const field = target === 'pdf' ? 'peakRssMb' : 'retainedMb';
    for (const now of report[target] ?? []) {
      const then = before[target]?.find((r) => r.document === now.document);
      if (!then) continue;
      for (const stage of now.stages) {
        const old = then.stages.find((s) => s.stage === stage.stage);
        if (!old) continue;
        const delta = Math.round((stage[field] - old[field]) * 100) / 100;
        const pct = old[field] ? Math.round((delta / old[field]) * 100) : 0;
        console.log(
          `${target.padEnd(8)}${now.document.padEnd(26)}${stage.stage.padEnd(16)}${String(old[field]).padStart(9)} → ${String(stage[field]).padStart(9)} MB  (${delta >= 0 ? '+' : ''}${delta} MB, ${pct >= 0 ? '+' : ''}${pct}%)${target === 'pdf' ? ' peak RSS' : ''}`
        );
      }
      if (target === 'pdf' && then.pdfSha256 !== now.pdfSha256) {
        console.log(`${target.padEnd(8)}${now.document.padEnd(26)}PDF bytes changed`);
      }
    }
  }
}
