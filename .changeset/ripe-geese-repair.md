---
"@marcusok/excel-exporter": patch
---

Fix several robustness details across export routes:

- Validate that every sheet entry is a non-null object up front (`each sheet must be an object`), so structurally malformed input no longer falls into the export routes with a bare TypeError.
- Isolate the browser download trigger from the fallback chain: if `triggerDownload` throws (e.g. a mocked/blocked DOM), the export no longer rebuilds via main-thread retry or stream fallback — the Blob is still returned and a warning is logged.
- Render `null`/`undefined` data rows as empty rows in both workbook and stream paths instead of throwing, including number/date `FormatSpec` columns.
- Skip empty cells in the fast-xlsx stream writer (no empty-string interning, no empty `<c>` elements), slightly reducing output size.
- Skip empty `border: {}` style objects instead of emitting a no-op border definition.
- `echartsToSheet`: throw a clear error when `categoryHeader` collides with the internal `__series_N` keys in wide layout.
- Add the missing `phase`/`duration` fields to the `WorkerResponse` type (the worker already sent them at runtime).
