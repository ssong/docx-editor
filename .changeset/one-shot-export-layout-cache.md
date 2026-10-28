---
'@docx-editor.dev/core': patch
'@docx-editor.dev/docx-to-markdown': patch
---

The headless page-count provider (`createAutomationPaginationProvider`) and `exportMarkdown` on a live `HeadlessDocumentView` now lay out with `reuseAcrossRevisions: false`. They dispose the session after one layout, so they no longer build state that is only useful for a later revision. Output is the same.
