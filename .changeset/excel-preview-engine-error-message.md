---
"@marcusok/excel-preview": patch
---

Keep the real message when an engine error arrives as a plain object with a `message` field (a future engine code, or any path that does not wrap it in `Error`). It previously stringified to `[object Object]`, discarding the only diagnostic detail.
