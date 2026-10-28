// Playwright config for the screenshot diff gate. See `e2e/screenshot-diff/README.md`.
//
// The spec compares each captured page with the PNG of the same name under
// SCREENSHOT_DIFF_BASELINE. `e2e/screenshot-diff/run.mjs` first fills that directory from the
// base engine (`--update-snapshots=all`), then runs again against the head engine to
// compare. SCREENSHOT_DIFF_ENGINE picks the checkout whose built packages the page uses.

import { defineConfig, devices } from '@playwright/test';
import { resolve } from 'node:path';

const PORT = Number(process.env.SCREENSHOT_DIFF_PORT ?? 5291);
const head = resolve(import.meta.dirname, '..');
const engine = resolve(process.env.SCREENSHOT_DIFF_ENGINE ?? head);
const baseline = resolve(
  process.env.SCREENSHOT_DIFF_BASELINE ?? resolve(head, 'local/screenshot-diff/baseline')
);
const results = resolve(
  process.env.SCREENSHOT_DIFF_RESULTS ?? resolve(head, 'local/screenshot-diff/results')
);
const chromium = process.env.SCREENSHOT_DIFF_CHROMIUM;
// Exact by default: any changed pixel fails. Anti-aliasing is deterministic for one Chromium
// build, and both sides of a comparison run in the same browser on the same machine. A change
// that moves text by design (for example rounding published coordinates) can loosen these
// for its own run: `threshold` is the per-pixel colour distance (0 to 1) that still counts
// as equal, `maxDiffPixels` the number of differing pixels a page may have.
const threshold = Number(process.env.SCREENSHOT_DIFF_THRESHOLD ?? 0);
const maxDiffPixels = Number(process.env.SCREENSHOT_DIFF_MAX_DIFF_PIXELS ?? 0);

export default defineConfig({
  testDir: '.',
  testMatch: 'screenshot-diff.spec.ts',
  snapshotPathTemplate: `${baseline}/{arg}{ext}`,
  outputDir: results,
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  // A retry would hide exactly the instability this gate must report.
  retries: 0,
  workers: 2,
  timeout: 120_000,
  expect: { toMatchSnapshot: { threshold, maxDiffPixels } },
  reporter: [['list'], ['html', { open: 'never', outputFolder: resolve(results, '../report') }]],
  use: {
    ...devices['Desktop Chrome'],
    viewport: { width: 1280, height: 1400 },
    deviceScaleFactor: 1,
    colorScheme: 'light',
    locale: 'en-US',
    timezoneId: 'UTC',
    reducedMotion: 'reduce',
    baseURL: `http://127.0.0.1:${PORT}`,
    ...(chromium ? { launchOptions: { executablePath: chromium } } : {}),
  },
  webServer: {
    command: `node e2e/screenshot-diff/serve.mjs --root "${engine}" --port ${PORT}`,
    cwd: head,
    url: `http://127.0.0.1:${PORT}/`,
    // Never reuse: a server left from the other side would serve the wrong engine.
    reuseExistingServer: false,
    timeout: 180_000,
    stdout: 'ignore',
  },
});
