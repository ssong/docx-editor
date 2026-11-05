#!/usr/bin/env node
// Fingerprints the layout of every fixture document with one checkout's built engine.
//
// Each document runs in its own Node process (`probe.mjs`), several at a time. The output
// directory holds one JSON file per document and measurer, plus `manifest.json`. Compare
// two output directories with `compare.mjs`. See scripts/layout-fingerprint/README.md.

import { spawn, execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { availableParallelism } from 'node:os';
import { join, relative, resolve } from 'node:path';
import { parseArgs } from 'node:util';
import { gzipSync } from 'node:zlib';

const here = import.meta.dirname;
const repo = resolve(here, '../..');

const { values } = parseArgs({
  options: {
    root: { type: 'string', default: repo },
    out: { type: 'string' },
    measurer: { type: 'string', default: 'fixed,shaped' },
    fixtures: { type: 'string', default: 'e2e/fixtures' },
    sizes: { type: 'string', default: '10,100,1000' },
    filter: { type: 'string' },
    jobs: { type: 'string', default: String(Math.max(1, Math.min(availableParallelism(), 8))) },
    timeout: { type: 'string', default: '300' },
    help: { type: 'boolean', default: false },
  },
});

if (values.help || !values.out) {
  console.log(`Usage: node scripts/layout-fingerprint/run.mjs --out <dir> [options]

Options:
  --root <dir>        checkout whose built packages lay out the documents (default: this one)
  --out <dir>         output directory; it is emptied first
  --measurer <list>   fixed, shaped, or both comma-separated (default: fixed,shaped)
  --fixtures <dir>    documents to lay out, searched recursively (default: e2e/fixtures)
  --sizes <a,b,...>   synthetic sizes from benchmarks/memory/fixtures.mjs ('' for none)
                      (default: 10,100,1000)
  --filter <text>     only documents whose name contains this text
  --jobs <n>          documents laid out at the same time (default: CPU count, at most 8)
  --timeout <s>       seconds before one document counts as failed (default: 300)

Build the packages of --root first: bun run --filter '@docx-editor.dev/i18n' build &&
bun run --filter '@docx-editor.dev/core' build && bun run --filter '@docx-editor.dev/fonts' build`);
  process.exit(values.help ? 0 : 2);
}

const root = resolve(values.root);
for (const entry of ['core/dist/export.js', 'core/dist/layout.js', 'fonts/dist/index.js']) {
  if (!existsSync(join(root, 'packages', entry))) {
    throw new Error(`${join(root, 'packages', entry)} is missing. Build the packages first.`);
  }
}
const measurers = values.measurer.split(',').filter(Boolean);
if (!measurers.every((m) => m === 'fixed' || m === 'shaped')) {
  throw new Error('--measurer takes fixed, shaped, or both');
}

// Documents come from THIS checkout, never from --root: both sides of a comparison must lay
// out the same bytes, including fixtures that only the head adds.
const fixtureRoot = resolve(repo, values.fixtures);
const documents = [];
(function collect(dir) {
  for (const name of readdirSync(dir).sort()) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) collect(path);
    else if (name.endsWith('.docx')) {
      const weight = statSync(path).size;
      documents.push({ name: relative(fixtureRoot, path), args: [path], weight });
    }
  }
})(fixtureRoot);
for (const units of values.sizes.split(',').filter(Boolean).map(Number)) {
  // About 20 KB of DOCX per unit, so the weight sorts it among the real fixtures.
  documents.push({
    name: `synthetic-${units}`,
    args: ['--synthetic', String(units)],
    weight: units * 20_000,
  });
}
const selected = values.filter
  ? documents.filter((d) => d.name.includes(values.filter))
  : documents;

const out = resolve(values.out);
rmSync(out, { recursive: true, force: true });
for (const measurer of measurers) mkdirSync(join(out, measurer), { recursive: true });

// Gzipped: the 1,000-page synthetic document alone is about 25 MB of JSON per measurer.
const fileName = (name) => `${name.replaceAll('/', '__')}.json.gz`;
const probe = join(here, 'probe.mjs');
const timeoutMs = Number(values.timeout) * 1000;

function layOut(document, measurer) {
  return new Promise((done) => {
    const started = performance.now();
    const child = spawn(
      process.execPath,
      [
        '--max-old-space-size=4096',
        probe,
        '--root',
        root,
        '--measurer',
        measurer,
        ...document.args,
      ],
      { stdio: ['ignore', 'pipe', 'pipe'] }
    );
    const stdout = [];
    let stderr = '';
    child.stdout.on('data', (chunk) => stdout.push(chunk));
    child.stderr.on('data', (chunk) => (stderr += chunk));
    const timer = setTimeout(() => child.kill('SIGKILL'), timeoutMs);
    child.on('close', (code, signal) => {
      clearTimeout(timer);
      const ms = Math.round(performance.now() - started);
      let text = Buffer.concat(stdout).toString('utf8').trim();
      if (code !== 0 || !text) {
        // The message, not the stack: a stack names file paths that differ between the
        // base and head checkouts, and the comparison must not report those.
        const reason = signal === 'SIGKILL' ? 'timed out' : firstErrorLine(stderr);
        text = JSON.stringify({ error: `layout failed: ${reason}` });
      }
      writeFileSync(join(out, measurer, fileName(document.name)), gzipSync(`${text}\n`));
      const parsed = JSON.parse(text);
      done({
        name: document.name,
        measurer,
        file: `${measurer}/${fileName(document.name)}`,
        ms,
        ...(parsed.error ? { error: parsed.error } : { pageCount: parsed.pageCount }),
      });
    });
  });
}

function firstErrorLine(stderr) {
  const line = stderr.split('\n').find((l) => /\b\w*Error\b/.test(l)) ?? stderr.split('\n')[0];
  return line?.trim() || 'no output';
}

const tasks = measurers.flatMap((measurer) => selected.map((document) => ({ document, measurer })));
// Largest work first, so a 1,000-page document does not start last and set the wall time.
tasks.sort((a, b) => b.document.weight - a.document.weight);
const results = [];
let next = 0;
let finished = 0;
const started = performance.now();
async function worker() {
  while (next < tasks.length) {
    const { document, measurer } = tasks[next++];
    const result = await layOut(document, measurer);
    results.push(result);
    finished += 1;
    const status = result.error ? `ERROR ${result.error}` : `${result.pageCount} pages`;
    console.log(
      `[${finished}/${tasks.length}] ${measurer} ${document.name}: ${status} (${result.ms} ms)`
    );
  }
}
await Promise.all(Array.from({ length: Number(values.jobs) }, worker));

let commit = null;
try {
  commit = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim();
} catch {
  // Not a git checkout; the manifest still names the root.
}
results.sort((a, b) => a.measurer.localeCompare(b.measurer) || a.name.localeCompare(b.name));
writeFileSync(
  join(out, 'manifest.json'),
  `${JSON.stringify({ format: 1, root, commit, measurers, node: process.version, documents: results }, null, 2)}\n`
);
const failed = results.filter((r) => r.error).length;
console.log(
  `Laid out ${results.length} documents in ${Math.round((performance.now() - started) / 1000)} s` +
    `${failed ? `, ${failed} failed (recorded, compared like any other result)` : ''}. Output: ${out}`
);
