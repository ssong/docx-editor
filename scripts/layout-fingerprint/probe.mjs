// Lays out one document and prints its layout fingerprint as one line of JSON.
//
// `run.mjs` starts this file once per document, so no document inherits caches or module
// state from another. That keeps a fingerprint independent of the order and the grouping of
// documents, which is what lets two runs on the same commit match byte for byte.
//
// The engine comes from `--root`, a checkout whose packages are built. The comparison in CI
// runs this same file against the base checkout and the head checkout, so the fingerprint
// format never differs between the two sides.
//
// Usage: node scripts/layout-fingerprint/probe.mjs --root <checkout> --measurer fixed|shaped
//        (<document.docx> | --synthetic <units>)

import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { syntheticDocx } from '../../benchmarks/memory/fixtures.mjs';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    root: { type: 'string' },
    measurer: { type: 'string', default: 'fixed' },
    synthetic: { type: 'string' },
    'export-options': { type: 'string' },
  },
});
const root = resolve(values.root ?? join(import.meta.dirname, '../..'));
const dist = (pkg, entry) => pathToFileURL(join(root, 'packages', pkg, 'dist', entry)).href;

const exporter = await import(dist('core', 'export.js'));
const layoutApi = await import(dist('core', 'layout.js'));

const bytes = values.synthetic
  ? syntheticDocx(Number(values.synthetic))
  : new Uint8Array(readFileSync(positionals[0]));

// Extra session options, such as a layout option under test. An engine that does not know
// an option ignores it.
const extraOptions = values['export-options'] ? JSON.parse(values['export-options']) : {};

async function openSession() {
  if (values.measurer === 'fixed') {
    // The engine's fixed measurer: no fonts, so the result depends on layout code alone.
    return exporter.openDocumentForExport(bytes, { ...extraOptions, reuseAcrossRevisions: false });
  }
  if (values.measurer !== 'shaped') throw new Error('--measurer takes fixed or shaped');
  // HarfBuzz over the packaged metric-compatible faces, the same origin the PDF exporter
  // puts after installed fonts. No installed fonts, so every machine shapes the same faces.
  const fonts = await import(dist('fonts', 'index.js'));
  const packaged = fonts.packagedFonts({
    install: false,
    fetcher: exporter.createPackagedFileFetch({
      trustedRoot: new URL('./', fonts.FONT_ASSET_ROOT),
      maxBytes: layoutApi.HARD_MAX_FONT_BYTES,
    }),
  });
  return exporter.openFontBackedDocumentForExport(bytes, { ...extraOptions, fonts: packaged });
}

const opened = await openSession();
if (!opened.ok) {
  console.log(JSON.stringify({ error: `open failed: ${opened.reason}` }));
  process.exit(0);
}
const { session } = opened;
const layout = await session.layout();

// Absolute span boxes come from the engine's own traversal. A change to how a span stores
// its box (for example relative to its line) must keep `absoluteBox` the same, so reading
// it here, not `span.box`, is what makes such a change comparable.
const spansByLine = new Map();
layoutApi.forEachSemanticSpan(layout, (visit) => {
  const { span, absoluteBox: box, page } = visit;
  const entry = [span.text, span.range.start, box.x - page.box.x, box.y - page.box.y, box.width];
  if (visit.paragraphId !== visit.line.range.paragraphId) entry.push(visit.paragraphId);
  let list = spansByLine.get(visit.line);
  if (!list) spansByLine.set(visit.line, (list = []));
  list.push(entry);
});

const pages = layout.pages.map((page) => ({
  w: page.box.width,
  h: page.box.height,
  lines: [],
}));
let orphanSpans = 0;
layoutApi.forEachSemanticStory(layout, (story) => {
  const target = pages[story.page.index];
  // Page-relative positions: a page that grows must not move every later page's numbers.
  const origin = story.page.box;
  layoutApi.forEachStoryParagraphFragment(
    story.host,
    (fragment, context) => {
      const kind = context.textboxDepth === 0 ? story.story : `textbox${context.textboxDepth}`;
      for (const line of fragment.lines) {
        const record = {
          s: kind,
          p: line.range.paragraphId,
          r: [line.range.start, line.range.end],
          x: context.storyOrigin.x + line.contentX - origin.x,
          y: context.storyOrigin.y + line.box.y + line.baseline - origin.y,
          w: line.box.width,
          h: line.box.height,
          spans: spansByLine.get(line) ?? [],
        };
        spansByLine.delete(line);
        if (line.drawings?.length) {
          record.drawings = line.drawings.map((d) => [d.start, d.x, d.y, d.width, d.height]);
        }
        target.lines.push(record);
      }
    },
    story.origin
  );
});
for (const list of spansByLine.values()) orphanSpans += list.length;

session.dispose();

const result = {
  sha256: createHash('sha256').update(bytes).digest('hex'),
  pageCount: layout.pages.length,
  pages,
};
// Spans the line walk did not reach. Zero on every fixture today; a change here means the
// two traversals stopped agreeing, which is itself worth a look.
if (orphanSpans > 0) result.orphanSpans = orphanSpans;
console.log(JSON.stringify(result));
