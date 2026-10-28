// The web server for the screenshot diff: builds the harness page against one checkout's
// engine, writes the fixture documents, and serves both.
//
// Usage: node e2e/screenshot-diff/serve.mjs --root <checkout> --port <port>
//
// The page source always comes from THIS checkout (the head), so both sides of a
// comparison run the same harness. It is copied into `<root>/local/screenshot-diff/` and
// built there, so `@docx-editor.dev/core` and `@docx-editor.dev/fonts` resolve to the
// built packages of `<root>`. Build `i18n`, `core` and `fonts` in `<root>` first.

import { execFileSync } from 'node:child_process';
import { cpSync, existsSync, mkdirSync, readFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { extname, join, resolve } from 'node:path';
import { parseArgs } from 'node:util';
import { SCREENSHOT_FIXTURES, fixtureBytes } from './fixtures.mjs';

const here = import.meta.dirname;
const head = resolve(here, '../..');
const { values } = parseArgs({
  options: {
    root: { type: 'string', default: head },
    port: { type: 'string', default: '5291' },
  },
});
const root = resolve(values.root);
for (const pkg of ['core', 'fonts']) {
  if (!existsSync(join(root, 'packages', pkg, 'dist'))) {
    throw new Error(`${root}/packages/${pkg}/dist is missing. Build the packages first.`);
  }
}

const work = join(root, 'local', 'screenshot-diff');
const pageSrc = join(work, 'page-src');
const pageDir = join(work, 'page');
mkdirSync(work, { recursive: true });
cpSync(join(here, 'page'), pageSrc, { recursive: true });
const vite = [join(root, 'node_modules/.bin/vite'), join(head, 'node_modules/.bin/vite')].find(
  existsSync
);
execFileSync(vite, ['build', '--config', join(pageSrc, 'vite.config.mjs'), '--outDir', pageDir], {
  cwd: root,
  stdio: 'inherit',
});

const documents = new Map();
for (const entry of SCREENSHOT_FIXTURES) {
  documents.set(`/fixtures/${entry.name}.docx`, await fixtureBytes(entry));
}

const types = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.wasm': 'application/wasm',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.woff2': 'font/woff2',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};
const server = createServer((req, res) => {
  const url = decodeURIComponent((req.url ?? '/').split('?')[0]);
  const type = types[extname(url)] ?? 'application/octet-stream';
  const doc = documents.get(url);
  if (doc) return res.writeHead(200, { 'Content-Type': type }).end(doc);
  const path = join(pageDir, url === '/' ? 'index.html' : url);
  if (!path.startsWith(pageDir)) return res.writeHead(403).end();
  try {
    const body = readFileSync(path);
    res.writeHead(200, { 'Content-Type': types[extname(path)] ?? type }).end(body);
  } catch {
    res.writeHead(404).end();
  }
});
server.listen(Number(values.port), '127.0.0.1', () => {
  console.log(`screenshot diff page for ${root} on http://127.0.0.1:${values.port}/`);
});
