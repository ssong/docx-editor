# Layout equivalence check

`bun run layout:check` proves that a change does not move text. It lays out every document in `e2e/fixtures` and three synthetic documents with a base commit and with your checkout. Then it compares a fingerprint of each layout and prints the differences.

Use it for changes that must keep the layout the same, for example memory work on the layout records. A change that fixes layout is expected to fail the check. CI shows which documents moved, so a reviewer can confirm that only the intended documents changed.

## What the fingerprint holds

For each document and each text measurer:

- The page count and the size of each page.
- For each line, in published order: the story (`body`, `header`, `footer`, `footnote`, `endnote`, or `textbox<depth>`), the paragraph id, the character range, the content x, the baseline y, the line box width and height, and the inline drawings.
- For each span on a line: the text, the start offset, and the x, y, and width of its box.

Positions are in points, relative to the top-left corner of the page. Span boxes come from `forEachSemanticSpan`, so a change to how a span stores its box must keep `absoluteBox` the same.

The check uses two text measurers:

- `fixed`: the engine's fixed measurer. It uses no fonts, so the result depends only on the layout code.
- `shaped`: HarfBuzz over the packaged fonts from `@docx-editor.dev/fonts`. It uses no installed fonts, so every machine gets the same result.

The synthetic documents come from `benchmarks/memory/fixtures.mjs`, at 10, 100, and 1,000 units (about one page each).

## Run the check

1. Build the packages in your checkout:

   ```bash
   bun run --filter '@docx-editor.dev/i18n' build
   bun run --filter '@docx-editor.dev/core' build
   bun run --filter '@docx-editor.dev/fonts' build
   ```

   Or add `--build-head` to the command in the next step.

1. Run the check against the base commit:

   ```bash
   bun run layout:check -- --base origin/main
   ```

   The script makes a Git worktree of the base commit in the system temp directory, installs and builds it, and keeps it for the next run. It builds the base again only when the base commit changes. Each document runs in its own Node.js process, so the result does not depend on the order of the documents. On 4 CPUs, one side takes about 2 minutes.

1. Read the report. Exit code 0 means the layout matches. Exit code 1 means it differs.

   ```text
   Layout differs in 1 of 214 documents (exact positions, exact spans)

   ✗ fixed demo.docx: 1 line differs
       page 3, body line /word/document.xml#0.0.19 [0, 33] "Tables in Word can vary from the ": baseline y 316.7 → 316.71
   ```

   When a line breaks at a different place, the report shows that line and stops for that document, because all later lines move too.

## Options

| Option | Default | What it does |
| --- | --- | --- |
| `--base <ref>` | `origin/main` | The commit to compare against. |
| `--tolerance <pt>` | `0` | The allowed difference for positions and widths, in points. `0` means exact. |
| `--spans <mode>` | `exact` | `exact` needs the same spans on each line. `merged` lets one head span join adjacent base spans of the same line, when the text, the start offset, the start x, and the end x match. |
| `--measurer <list>` | `fixed,shaped` | One measurer or both. |
| `--sizes <list>` | `10,100,1000` | Synthetic document sizes. `''` for none. |
| `--filter <text>` |  | Only documents whose name contains this text. |
| `--jobs <n>` | CPU count, at most 8 | Documents laid out at the same time. |
| `--work <dir>` | system temp directory | Where the base worktree and the two outputs go. |
| `--build-head` | off | Build the packages of your checkout first. |
| `--markdown <file>` |  | Also write the report as Markdown. |

For example, a change that rounds published coordinates to 1/100 pt:

```bash
bun run layout:check -- --tolerance 0.01
```

A change that joins the spans of a run segment:

```bash
bun run layout:check -- --spans merged --tolerance 0.01
```

## The parts

- `probe.mjs` lays out one document with the engine of a given checkout and prints its fingerprint.
- `run.mjs` (`bun run layout:fingerprint`) runs `probe.mjs` for every document and writes a directory of gzipped fingerprints and `manifest.json`.
- `compare.mjs` (`bun run layout:compare -- <base-dir> <head-dir>`) compares two such directories.
- `check.mjs` (`bun run layout:check`) prepares the base worktree and runs the other three.

The scripts of your checkout lay out both sides. So the base and the head always use the same fingerprint format and the same documents, also when the base commit has no `scripts/layout-fingerprint`.

## In CI

The `layout equivalence` job in `.github/workflows/ci.yml` runs on pull requests that change `packages/`, the fixtures, or these scripts. It compares against the merge base with the target branch, in exact mode, and writes the report to the job summary. When the layout differs, it uploads both fingerprint directories as the `layout-fingerprints` artifact.

A pull request that changes the layout on purpose gets the `layout-change-expected` label. Then the job passes and still shows the report. Add the label, then run the job again.
