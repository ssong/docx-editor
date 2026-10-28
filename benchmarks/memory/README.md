# Core memory benchmark

`bun run bench:memory` measures how much memory `@docx-editor.dev/core` uses for a range of document sizes, on a server (Node.js) and in a browser (Chromium). It also measures the peak memory of a PDF export with `@docx-editor.dev/docx-to-pdf`. It writes a JSON report to `benchmarks/memory/results/`, so you can compare a change against an earlier run.

The benchmark answers two questions:

- How much memory does an open document keep alive, and how does that grow with the document?
- Does the engine give the memory back when you close the document?

## Run the benchmark

1. Build the packages that the benchmark loads:

   ```bash
   bun run --filter '@docx-editor.dev/i18n' build
   bun run --filter '@docx-editor.dev/core' build
   bun run --filter '@docx-editor.dev/fonts' build
   ```

   For the PDF export half, build the PDF package and the packages it loads:

   ```bash
   bun run build:pdf
   ```

1. Run the benchmark:

   ```bash
   bun run bench:memory
   ```

   A full run takes about 30 minutes, because the browser half opens and edits documents of up to 1,000 pages three times each. For a quick check, run fewer sizes:

   ```bash
   bun run bench:memory -- --sizes 10,100 --real '' --runs 1
   ```

1. To compare against an earlier report, pass it with `--compare`:

   ```bash
   bun run bench:memory -- --compare benchmarks/memory/results/<earlier-report>.json
   ```

If Playwright cannot find its browser, pass a Chromium executable with `--chromium <path>` or the `MEMORY_BENCH_CHROMIUM` environment variable.

## Options

| Option | Default | Description |
| --- | --- | --- |
| `--sizes` | `10,50,100,250,500,1000` | Synthetic document sizes, in content units. One unit is about one page. |
| `--real` | `e2e/fixtures/demo.docx,e2e/fixtures/typing-perf-521pp.docx` | Real documents to add. Pass `''` for none. |
| `--runs` | `3` | Repetitions for each document. The report keeps the median of each value. |
| `--target` | `all` | `all`, `server`, `browser`, or `pdf`. `all` includes `pdf` when `@docx-editor.dev/docx-to-pdf` is built. |
| `--server-edits` | `200` | Single-character edits in the server edit stage. |
| `--browser-edits` | `50` | Typed characters in the browser edit stage. |
| `--fonts` | `shaped` | `shaped` loads the packaged default fonts in the browser, as production hosts do. `fixed` uses the fixed-width measurer. |
| `--out` | `results/<date>-<commit>.json` | The report file. |
| `--compare` |  | An earlier report. The tool prints the change for each stage. |
| `--no-build` |  | Reuse the benchmark page from the previous run. |
| `--chromium` |  | A Chromium executable. |

## Documents

The synthetic documents come from `fixtures.mjs`. Each content unit has a heading, five body paragraphs with bold, italic, and colored runs, a three-item bullet list, and a 4×5 table on every second unit. Every page has a header, and a footer with a `PAGE` field. The generator uses fixed zip timestamps, so a size always produces the same bytes. The report records the SHA-256 of each document.

The two real documents are repository fixtures: `demo.docx` (images, tables, and review marks) and `typing-perf-521pp.docx` (521 pages, 81 sections).

## What each half measures

Both halves force garbage collection before each reading. A reading is the memory that is still reachable, not the memory that a stage used for a moment.

### Server

`server-probe.mjs` runs once for each document and repetition, in a new Node.js process with `--expose-gc`. It imports the built package, as a Node.js service does.

| Stage | What it does |
| --- | --- |
| `modules loaded` | Imports `@docx-editor.dev/core/store` and `@docx-editor.dev/core/export`. This is the baseline. |
| `open` | `openHeadlessDocument(bytes)`: parse, canonical tree, and store. |
| `layout` | `openDocumentForExport(view, { reuseAcrossRevisions: false }).layout()` with the fixed measurer: the PDF export and page-count path, which lay out once and dispose the session. |
| `release` | Disposes the export session and drops the view. Memory that stays is held by module-level state. |
| `reopen store` | `readOoxmlPackage` and a new `TreePackageStore`, for the edit stage. |
| `edit xN` | N single-character `insertText` transactions. Each one is an undo step. |
| `save` | `writeOoxmlPackage` on the edited package. |

Each stage reports `heapUsedMb` (the V8 heap after collection), `retainedMb` (the change from `modules loaded`), `deltaMb` (the change from the previous stage), `externalMb`, `rssMb`, and `peakRssMb` (the process high-water mark, which includes short-lived garbage).

### PDF export

`pdf-probe.mjs` runs once for each document and repetition, in a new Node.js process with `--expose-gc`. It imports the built `@docx-editor.dev/docx-to-pdf` package and converts the document with `exportPdf`, with only the packaged fonts, so the result does not depend on the fonts that the machine has.

| Stage            | What it does                                                   |
| ---------------- | -------------------------------------------------------------- |
| `modules loaded` | Imports `@docx-editor.dev/docx-to-pdf`. This is the baseline.  |
| `export`         | `exportPdf(bytes)`: open, layout, paint, and save in one call. |
| `release`        | Drops the result.                                              |

The stages report the same fields as the server half. The value to watch is `peakRssMb` after `export`: the process high-water mark of the whole conversion. The export frees most of its memory before it returns, so `retainedMb` shows little. The report also keeps `layoutMs` and `paintMs` from the export timings, the size of the PDF, and its SHA-256 (`pdfSha256`). The benchmark stops if two runs of one document give different bytes, and `--compare` prints `PDF bytes changed` when the bytes differ from the earlier report.

### Browser

The browser half builds `browser/` with Vite in production mode. The page mounts the core engine with `createDocxEditor`, without a framework adapter, inside the same scroll container classes that `DocxEditor.Viewport` renders. The engine finds its scroller by the `docx-editor__scroll-container` class and builds only the pages near the viewport. Each repetition launches a new headless Chromium.

| Stage | What it does |
| --- | --- |
| `page loaded` | The page and the engine code, before any document. This is the baseline. |
| `fonts loaded` | `loadDefaultFonts()` from `@docx-editor.dev/fonts`. Only with `--fonts shaped`. |
| `open` | `createDocxEditor({ document, fonts })` and `attach`, until layout and paint stop. |
| `scroll through` | Scrolls from the first page to the last page, one viewport at a time. |
| `type xN` | Types one character at N paragraphs spread through the document. |
| `destroy` | `editor.destroy()`. Memory that stays is held by module-level state. |

Each stage reports `heapUsedMb` (the JavaScript heap), `retainedMb`, `deltaMb`, `pageTotalMb` and `domMb` (from `performance.measureUserAgentSpecificMemory()`, which also counts DOM objects and WebAssembly memory), `domNodes`, and `listeners`. With `--fonts shaped`, `retainedMb` counts from `fonts loaded`, because the host loads the fonts once for every document.

V8 counts compiled code in the heap, so `destroy` keeps a few megabytes more than the baseline even when the engine releases the whole document.

## Read the results

- Memory values are stable from run to run. A change of more than about 2% on a large document is a real change.
- Time values (`ms`) depend on the machine and its load. Compare them only between runs on the same machine.
- `release` on the server and `destroy` in the browser should come back close to the baseline. Memory that stays after them, and that grows with the document size, is held by a module-level cache.
- `scripts/memory/edit-memory.mjs` (`bun run memory:edits`) finds what holds memory that grows while you edit. Use it after this benchmark shows growth in an edit stage.

## Find what holds the memory

When the benchmark shows a change, heap snapshots show where the memory is:

- `server-snapshot.mjs` writes Node.js heap snapshots after `open`, `layout`, and `release`, and before and after 200 edits.
- `browser-snapshot.mjs` builds the page from the engine source without minification, and writes Chromium heap snapshots after `open`, after typing, and after `destroy`.
- `analyze-snapshot.mjs` summarizes a snapshot: object counts and sizes by constructor, the memory each module or closure variable keeps alive, and the memory held through each object property.

The built package is minified, so a snapshot of it has short names. To get the source names on the server, bundle the script from source first. The root `tsconfig.json` maps the package to its source:

```bash
bun build benchmarks/memory/server-snapshot.mjs --target=node --outdir=local/memory-bench/snapshot-build
node --expose-gc local/memory-bench/snapshot-build/server-snapshot.js <document.docx> local/memory-bench/snapshots
node benchmarks/memory/analyze-snapshot.mjs local/memory-bench/snapshots/2-layout.heapsnapshot
```

```bash
node benchmarks/memory/browser-snapshot.mjs <document.docx> local/memory-bench/browser-snapshots
```

## Results

`report.mjs` renders a report as Markdown tables, optionally with the change from an earlier report:

```bash
node benchmarks/memory/report.mjs benchmarks/memory/results/<report>.json --compare <earlier>.json
```

Commit a report from `results/` when it records a baseline or the effect of a memory change. Name the machine in the pull request, because absolute values change a little between Chromium and Node.js versions.
