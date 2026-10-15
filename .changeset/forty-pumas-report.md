---
"@marcusok/excel-exporter": minor
---

`ExportResult.engine` now reports `"fast-xlsx"` on stream-route exports instead of `"modern-xlsx"` — the fast stream is the package's own pure-JS writer, not the WASM engine. The field type widens to `"modern-xlsx" | "fast-xlsx"`; use it (not `mode`) to tell engines apart. Also, `terminateWorker()` now rejects in-flight requests with the same no-fallback policy as the worker timeout (no main-thread rerun of the same large build; retry recreates the worker).
