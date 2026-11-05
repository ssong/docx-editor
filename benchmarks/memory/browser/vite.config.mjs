import { resolve } from 'node:path';
import { defineConfig } from 'vite';

const root = resolve(import.meta.dirname, '../../..');
// MEMORY_BENCH_SOURCE=1 builds the engine from source without minification, so names in a
// heap snapshot match the code. The benchmark itself measures the built package.
const fromSource = process.env.MEMORY_BENCH_SOURCE === '1';

// A production build of the benchmark page, so the measured heap matches shipped code:
// minified, no dev server client, no source maps. The root workspace does not depend on
// `@docx-editor.dev/fonts`, so it resolves to the built package directly.
export default defineConfig({
  root: import.meta.dirname,
  base: './',
  logLevel: 'warn',
  resolve: {
    alias: [
      {
        find: '@docx-editor.dev/fonts',
        replacement: resolve(root, 'packages/fonts/dist/index.js'),
      },
      ...(fromSource
        ? [
            // Every entry, like the root tsconfig `paths`: one copy of the engine, all source.
            {
              find: /^@docx-editor\.dev\/core\/(?!styles\/)(.+)$/,
              replacement: `${resolve(root, 'packages/core/src')}/$1`,
            },
            {
              find: /^@docx-editor\.dev\/core$/,
              replacement: resolve(root, 'packages/core/src/index.ts'),
            },
          ]
        : []),
    ],
  },
  build: {
    sourcemap: false,
    emptyOutDir: true,
    chunkSizeWarningLimit: 10_000,
    ...(fromSource ? { minify: false } : {}),
  },
});
