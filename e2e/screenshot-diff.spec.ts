// The screenshot diff gate: renders the first pages of each fixture with the core engine
// and compares every pixel with the same page from the base engine. Run it through
// `bun run test:screenshots`, which renders the base first. See `e2e/screenshot-diff/README.md`.

import { expect, test } from '@playwright/test';
import { SCREENSHOT_FIXTURES } from './screenshot-diff/fixtures.mjs';

declare global {
  interface Window {
    __screenshotDiff: {
      open(url: string): Promise<{ pages: number }>;
      show(index: number): Promise<void>;
      destroy(): void;
    };
  }
}

for (const fixture of SCREENSHOT_FIXTURES) {
  test(fixture.name, async ({ page }) => {
    await page.goto('/');
    await page.waitForFunction(() => !!window.__screenshotDiff);
    const { pages } = await page.evaluate(
      (url) => window.__screenshotDiff.open(url),
      `/fixtures/${fixture.name}.docx`
    );
    // A page count change is a layout change even when the captured pages match.
    expect(`${pages}\n`).toMatchSnapshot([fixture.name, 'page-count.txt']);
    for (let index = 0; index < Math.min(pages, fixture.pages); index += 1) {
      await page.evaluate((i) => window.__screenshotDiff.show(i), index);
      const sheet = page.locator(`.docx-page[data-page-index="${index}"]`);
      const png = await sheet.screenshot({ animations: 'disabled', caret: 'hide' });
      // Soft, so one run reports every changed page of the document, not only the first.
      expect.soft(png).toMatchSnapshot([fixture.name, `page-${index + 1}.png`]);
    }
  });
}
