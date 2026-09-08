---
"@marcusok/excel-exporter": patch
---

- fix: a timed-out worker export (fixed 120s) no longer leaves the worker cached and running — the worker is terminated and dropped so the next export creates a fresh one, and sibling requests dispatched to the same worker are rejected immediately (their callers degrade to the main-thread retry) instead of queueing behind a wedged worker.
- fix: `format: { type: "number" }` columns now render `null`/`undefined` values as empty cells on every path. Previously `Number(null) === 0` silently turned missing values into `0` while `undefined` became `""` (asymmetric, and `0` is meaningful in financial data).
- fix: the SheetJS fallback warning now lists configured layout features it drops (`width` / `freezeRows` / `autoFilter`) alongside "styles stripped", matching the stream path's per-feature warnings instead of degrading them silently.
- fix: `configureWasm()` with a new `wasmUrl` after a successful load now warns that modern-xlsx's idempotent `initWasm` keeps the already-loaded module on that thread (the previous JSDoc/README claim that the next `ensureLoaded` "re-initializes from the new URL" was not what the real dependency does — verified against modern-xlsx 1.2.0's module-level `initialized` guard). Tests updated to describe the loader's actual state-machine contract.
- perf: the Workbook path deduplicates identical cell styles — the same `CellStyle` used by N header cells or columns now shares one style index instead of appending N duplicate font/fill/xf records (modern-xlsx's `StyleBuilder.build` never dedupes; 50 styled header cells previously produced 50 identical records).
- docs: `ColumnConfig.format` / `ExportResult.error` doc comments now state the `thousands` cross-path difference (Workbook renders `#,##0` via numFormat; stream/SheetJS keep the cell a number without separators), the null/undefined-as-empty-cell rule, and that `error` is also set on `success: true` fallback results (degradation reason).
- chore: the published tarball now includes `CHANGELOG.md`.
- tests: 94 → 100 (CI runs 96): worker timeout termination + sibling rejection, style dedup round-trip, fallback dropped-features warning, wasmUrl-change warning, number-spec null/undefined rendering.
