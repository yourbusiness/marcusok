---
"@marcusok/excel-exporter": patch
---

Render blank date/datetime FormatSpec values as empty cells on every path, isolate async callback rejections (including the overlay chain), preserve `AbortError` during the terminal stream fallback, validate null column/merge entries with actionable errors, and align the docs with the single-slot worker queue and main-thread-only WASM timeouts.
