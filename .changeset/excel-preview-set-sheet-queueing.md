---
"@marcusok/excel-preview": patch
---

Queue `setSheet()` calls made before the parse resolves and replay the last one once the preview boots. Previously such calls were silent no-ops (the documented workaround was to call from `onParsed`), and the package's own JSDoc example walked into exactly that trap. Invalid queued names are reported through `onError` when the replay runs, with the same wording and code as a runtime call. `getSheetNames()` still returns `[]` until the parse resolves (synchronous getter), and `destroy()` drops any queued call. The docs site (en + zh) and the JSDoc now describe the queueing behavior.
