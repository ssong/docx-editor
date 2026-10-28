// The browser half of the memory benchmark: the core engine alone, with no framework adapter.
//
// `run.mjs` builds this page with Vite in production mode and drives `window.__memoryBench`
// from Playwright. Each call resolves only after layout and paint settle, so a heap reading
// taken after it measures a finished state rather than work still in flight.

import '@docx-editor.dev/core/styles/editor.css';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { loadDefaultFonts } from '@docx-editor.dev/fonts';

const viewport = document.getElementById('viewport');
const host = document.getElementById('host');
let editor = null;
let fonts;

const frame = () => new Promise((done) => requestAnimationFrame(() => done()));
const sleep = (ms) => new Promise((done) => setTimeout(done, ms));

function pageCount() {
  return editor?.surface?.publishedLayout().pages.length ?? 0;
}

/**
 * Wait until the editor stops working: no open in progress, no pending layout, and the
 * published page count and painted height unchanged for half a second. Large documents
 * keep laying out and painting in background slices after `isOpening` turns false, so the
 * flags alone end too early. Resolves with the time of the last observed change.
 */
async function settled() {
  let last = null;
  let lastChange = performance.now();
  let quiet = 0;
  while (quiet < 5) {
    await sleep(100);
    const busy =
      editor?.snapshot().isOpening ||
      host.matches('[data-docx-layout-pending]') ||
      !!host.querySelector('[data-docx-layout-pending]');
    const state = `${pageCount()}:${viewport.scrollHeight}:${host.childElementCount}`;
    if (busy || state !== last) {
      last = state;
      lastChange = performance.now();
      quiet = 0;
    } else {
      quiet += 1;
    }
  }
  await frame();
  return lastChange;
}

window.__memoryBench = {
  /** Fetch the packaged default fonts, so their bytes are not counted as document cost. */
  async loadFonts() {
    fonts = await loadDefaultFonts();
  },

  async open(url, layoutSpans = 'word') {
    const bytes = new Uint8Array(await (await fetch(url)).arrayBuffer());
    const t0 = performance.now();
    editor = createDocxEditor({ document: bytes, layoutSpans, ...(fonts ? { fonts } : {}) });
    // The editor reports `document.open` once the last page of a sliced open is laid out.
    const laidOut = new Promise((done, fail) => {
      const off = editor.on('diagnostic', (event) => {
        if (event.kind !== 'document.open') return;
        off();
        if (event.outcome === 'opened') done();
        else fail(new Error(`open ${event.outcome}: ${event.code ?? ''}`));
      });
    });
    editor.attach(host);
    await laidOut;
    const end = await settled();
    return { ms: Math.round(end - t0), pages: pageCount() };
  },

  /** Scroll top to bottom one viewport at a time, then back to the top. */
  async scrollThrough() {
    const t0 = performance.now();
    for (let top = 0; top < viewport.scrollHeight; top += viewport.clientHeight) {
      viewport.scrollTop = top;
      await frame();
      await frame();
    }
    viewport.scrollTop = 0;
    const end = await settled();
    return { ms: Math.round(end - t0) };
  },

  /** Type one character at the start of `count` paragraphs spread through the document. */
  async type(count) {
    const surface = editor.surface;
    const t0 = performance.now();
    for (let step = 0; step < count; step += 1) {
      const ids = surface.session.paragraphIds();
      const id = ids[Math.floor((step / count) * ids.length)];
      const at = { paragraphId: id, offset: 0 };
      surface.setSelection({ anchor: at, head: at });
      surface.type('x');
      surface.flushPendingInput();
      await frame();
      await frame();
    }
    const end = await settled();
    return { ms: Math.round(end - t0) };
  },

  destroy() {
    editor.destroy();
    editor = null;
    host.replaceChildren();
  },
};
