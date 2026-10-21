---
"@marcusok/excel-exporter": patch
---

Update the shared wasm-loader's reconfiguration warnings (bundled from xlsx-core into this package's dist): the `wasmUrl` warning now also points at `terminateParseWorker()` from `@marcusok/excel-preview` for pages using both packages, and the misleading "recreate the preview" advice for `parseWorkerUrl` changes is gone.
