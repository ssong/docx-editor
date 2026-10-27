# Screenshot diff gate

The gate renders the first pages of a fixture set with two builds of the engine and compares every pixel. Use it for changes that must not change what the reader sees: memory work on layout records, span merging, coordinate rounding, and font loading.

```bash
bun run test:screenshots                       # head against origin/main
bun run test:screenshots -- --base <git ref>   # head against another commit
bun run test:screenshots -- --repeat           # head against head: the stability check
```

Add `--chromium /opt/pw-browsers/chromium` in a cloud container where Playwright's own browser is not installed.

On failure, open `local/screenshot-diff/run/report/index.html`. Each changed page shows the expected (base), actual (head) and diff images. CI runs the same comparison on every pull request (`.github/workflows/screenshot-diff.yml`) and uploads the report as the `screenshot-diff` artifact when it fails.

## How it works

1. `run.mjs` finds the base checkout. With `--base <ref>` it adds a worktree at `local/screenshot-diff/base/`, installs it, and builds `i18n`, `core` and `fonts` there. With `--base-root <dir>` it uses a checkout that is already built. It also builds the head packages unless you pass `--skip-build`.
2. Playwright runs `e2e/screenshot-diff.spec.ts` against the base engine with `--update-snapshots=all`. This writes the base PNGs to `<out>/base/`.
3. Playwright runs the same spec against the head engine. Each page must match its base PNG exactly, and the document page count must match too.

`serve.mjs` is the web server. It copies `page/` into the checkout under test (`<checkout>/local/screenshot-diff/`) and builds it there with Vite. So the page imports that checkout's built `@docx-editor.dev/core` and `@docx-editor.dev/fonts`, but the harness code is the same on both sides. Only the engine changes.

The page opens each document with `createDocxEditor` and `packagedFonts()`, like the demo, but with no framework adapter. It waits until the published layout, the page DOM and the font set stay unchanged for one second. It hides the engine's blinking caret. Playwright captures each `.docx-page` at 1280×1400, device scale factor 1, in light mode.

## Fixtures

`fixtures.mjs` lists the documents and the number of leading pages to capture:

- `demo.docx`, tables (`with-tables`, `repeated-table-header`), and images (`example-with-image`, `images-wrap-sides`)
- `comprehensive-word-element-test.docx`, which has CJK and Arabic text among many other elements
- `rtl-table-bidivisual.docx`, and two documents with many justified paragraphs
- `synthetic-scripts`: generated justified, distributed, centred and right-aligned Latin text, a dot-leader tab, Chinese, Japanese and Korean paragraphs (also justified), and Arabic and Hebrew paragraphs, including justified and mixed-direction text
- `synthetic-memory-bench`: four units of the memory benchmark document (`benchmarks/memory/fixtures.mjs`)

To add a case, add an entry. A `file` is a path under `e2e/fixtures/`. A `build` function returns DOCX bytes and must be deterministic.

## Tolerance

The default is exact: one changed pixel fails. Both sides run in the same Chromium on the same machine, and repeated runs on one commit match exactly. A 0.01 pt shift of every line moved pixels on 9 of the 11 documents, so the gate sees changes at that scale.

A change that moves text by design, for example rounding published coordinates, can loosen the comparison for its own run. Use `--threshold <0..1>` for the per-pixel colour distance that still counts as equal, and `--max-diff-pixels <n>` for the number of differing pixels a page may have. The same values are also available as the `SCREENSHOT_DIFF_THRESHOLD` and `SCREENSHOT_DIFF_MAX_DIFF_PIXELS` environment variables. Report the values you used in the pull request.

## Limits

- The baselines are not committed. Each run renders the base again, so a Chromium update or different system fonts do not break the gate. The cost is a second build and render, about one minute.
- CJK and some other scripts use the system fallback fonts of the machine, because `packagedFonts()` covers only Word's Latin defaults. Results are comparable on one machine, not across machines.
- The gate captures the first pages only. Use the layout equivalence gate for whole documents.
