---
'@docx-editor.dev/docx-to-pdf': patch
'@docx-editor.dev/core': minor
---

PDF export paints and compresses one page at a time, which lowers peak memory for long documents without changing the output bytes. `forEachSemanticSpan` accepts a `page` option to visit one page's spans.
