---
'@docx-editor.dev/core': minor
---

Add the experimental `compactLayoutBoxes` option to `createDocxEditor` and `openDocumentForExport`. With it, each text span box stores its coordinates as small integers, so layout uses less memory. Coordinates are rounded to 1/20000 pt. A compact box reads like any `LayoutBox`, but an object spread (`{ ...span.box }`) does not copy its coordinates. The option is off by default.
