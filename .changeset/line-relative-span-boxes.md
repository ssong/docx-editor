---
'@docx-editor.dev/core': minor
---

Laid-out lines share their text spans with the layout cache, so layout uses less memory. A span box's `y` is now relative to the top of its line; `spanStoryBox(line, span)` returns the box in story coordinates.
