---
"@marcusok/excel-exporter": patch
---

Progress-contract and documentation fixes from a full-repo audit:

- The stream path's final 1,000-row progress checkpoint is now skipped when it would land exactly on the last row: whole-thousand row counts previously emitted a duplicate `onProgress(1)` immediately before the terminal 1 that the `onProgress` contract promises fires exactly once. Regression tests added at both the engine and the entry level.
- `FormatSpec` pattern-token docs now state precisely that superset tokens are only partially verbatim on the stream path: `"mmm"` parses its `mm` prefix and emits a stray `m` (`"mmm"` -> `"09m"`), while the Workbook path renders the month abbreviation via numFormat.
- The `configureWasm` runtime warning and JSDoc now name the import location of `terminateWorker()` (`@marcusok/excel-exporter/worker-utils` subpath — it is not exported from the main entry, so the warning previously pointed at a function callers could not resolve).
- Internal cleanup, no behavior change: `buildWorksheetXml` parameters are now required (the optional declarations contradicted the non-null-asserted implementation and the single call site), and the duplicated `columnName` helper was consolidated into the shared `column-tree` export.
