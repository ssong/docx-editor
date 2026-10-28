---
'@docx-editor.dev/core': patch
---

Undo history uses less memory in long documents. An older undo step no longer keeps its own copy of the list of top-level blocks, so 200 edits in a 1,270-page document keep about 14 MB less.
