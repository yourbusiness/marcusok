---
"@marcusok/excel-exporter": patch
"@marcusok/excel-preview": patch
---

Ignore explicitly `undefined` options in `configureWasm` so an unset environment variable no longer clears an already-configured WASM/worker URL. The shared loader change ships with both Excel business packages that bundle it.
