---
'@docx-editor.dev/core': minor
'@docx-editor.dev/docx-to-markdown': patch
---

A new beta layout option joins the words of a run segment into one text span. Set `spans: 'run'` in `SemanticLayoutOptions` or `openDocumentForExport`, or `layoutSpans: 'run'` in `createDocxEditor`. Lines break at the same places and text stays at the same positions, but layout keeps fewer span records and paint creates fewer elements. Justified lines keep a span boundary at each stretched space. The default, `'word'`, keeps one span per word.
