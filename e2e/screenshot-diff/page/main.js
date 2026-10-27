// The browser half of the screenshot diff: the core engine alone, with no framework adapter.
//
// `serve.mjs` copies this page into the checkout under test and builds it there with Vite,
// so `@docx-editor.dev/core` and `@docx-editor.dev/fonts` resolve to THAT checkout's built
// packages. The page itself stays the same on both sides of a comparison; only the engine
// changes. `e2e/screenshot-diff.spec.ts` drives `window.__screenshotDiff` from Playwright.

import '@docx-editor.dev/core/styles/editor.css';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { packagedFonts } from '@docx-editor.dev/fonts';

const viewport = document.getElementById('viewport');
const host = document.getElementById('host');
let editor = null;
let layout = null;
let mutations = 0;

new MutationObserver((records) => {
  mutations += records.length;
}).observe(host, { subtree: true, childList: true, attributes: true, characterData: true });

const frame = () => new Promise((done) => requestAnimationFrame(() => done()));
const sleep = (ms) => new Promise((done) => setTimeout(done, ms));

function pageCount() {
  return editor?.surface?.publishedLayout().pages.length ?? 0;
}

/**
 * Wait until the editor stops working. Fonts resolve after the parse and lay the document
 * out again, and pages paint in background slices, so a pass counts as finished only when
 * the published layout, the painted DOM and the font set stay unchanged for a full second.
 */
async function settled() {
  let quiet = 0;
  let seen = mutations;
  while (quiet < 10) {
    await sleep(100);
    const current = editor?.surface?.publishedLayout() ?? null;
    const busy =
      editor?.snapshot().isOpening ||
      document.fonts.status === 'loading' ||
      host.matches('[data-docx-layout-pending]') ||
      !!host.querySelector('[data-docx-layout-pending]');
    if (busy || current !== layout || mutations !== seen) {
      layout = current;
      seen = mutations;
      quiet = 0;
    } else {
      quiet += 1;
    }
  }
  await document.fonts.ready;
  await frame();
  await frame();
}

window.__screenshotDiff = {
  async open(url) {
    const bytes = new Uint8Array(await (await fetch(url)).arrayBuffer());
    editor = createDocxEditor({ document: bytes, fonts: packagedFonts() });
    const opened = new Promise((done, fail) => {
      const off = editor.on('diagnostic', (event) => {
        if (event.kind !== 'document.open') return;
        off();
        if (event.outcome === 'opened') done();
        else fail(new Error(`open ${event.outcome}: ${event.code ?? ''}`));
      });
    });
    editor.attach(host);
    await opened;
    await settled();
    return { pages: pageCount() };
  },

  /** Scroll page `index` to the top of the viewport and wait for its paint to settle. */
  async show(index) {
    if (!editor.surface.revealPage(index)) throw new Error(`no page ${index}`);
    await settled();
    if (!host.querySelector(`.docx-page[data-page-index="${index}"]`)) {
      throw new Error(`page ${index} is not painted`);
    }
  },

  destroy() {
    editor?.destroy();
    editor = null;
    layout = null;
    host.replaceChildren();
  },
};
