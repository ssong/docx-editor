import { resolve } from 'node:path';
import { defineConfig } from 'vite';

// `serve.mjs` copies this directory to `<checkout>/local/screenshot-diff/page/` and builds
// it there, so the engine resolves from that checkout's `node_modules`. The root workspace
// does not depend on `@docx-editor.dev/fonts`, so it resolves to the built package directly.
const checkout = resolve(import.meta.dirname, '../../..');

export default defineConfig({
  root: import.meta.dirname,
  base: './',
  logLevel: 'warn',
  resolve: {
    alias: [
      {
        find: '@docx-editor.dev/fonts',
        replacement: resolve(checkout, 'packages/fonts/dist/index.js'),
      },
    ],
  },
  build: {
    sourcemap: false,
    emptyOutDir: true,
    chunkSizeWarningLimit: 10_000,
  },
});
