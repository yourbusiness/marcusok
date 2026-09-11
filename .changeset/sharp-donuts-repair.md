---
"@marcusok/excel-exporter": minor
---

Configurable worker timeout, input-validation hardening, and degradation-chain fixes:

- **feat: `configureWasm({ workerTimeoutMs })`** — the worker export timeout (previously a fixed 120s) is now configurable. A timed-out export still terminates the shared worker and rejects its sibling requests, so raise it only for legitimately huge exports.
- **fix: empty `sheets: []` no longer produces a corrupt file reported as success.** Previously the Workbook build threw, the export degraded to the fast stream and resolved `success: true` with a zero-sheet workbook Excel flags as corrupt (ECMA-376 requires at least one sheet). `exportExcel` pre-flight validation and `exportAsStream` now both reject with a clear `at least one sheet is required` error.
- **fix: a synchronous `postMessage` failure no longer kills the shared worker 120s later.** Un-cloneable payloads (e.g. Symbol/function values in row data) throw DataCloneError synchronously; the request previously stayed in the pending map with its timeout timer live, so the timer eventually fired and terminated the healthy shared worker, also rejecting every sibling request. The pending entry and timer are now cleaned up immediately.
- **fix: worker stream-route failures stop after one main-thread retry.** On the >= 50,000-row stream route the main-thread retry already runs the identical fast stream on identical input, so the previous third attempt via the terminal fallback failed deterministically (and logged a misleading "falling back" warning). The chain now fails directly with the combined reason.
- **fix: reusing the same `ColumnConfig` object in two places (a diamond, not a cycle) now throws.** It previously passed the path-based cycle check and was walked twice, silently emitting duplicate data columns.
- **fix: sheet names beginning or ending with an apostrophe are rejected** (Excel refuses such names), matching the other ECMA-376 name validations.
- **fix: custom date/time patterns interpret `mm` as minutes when followed by `ss`** (Excel's numFormat convention), so e.g. `mm:ss` renders `20:30` on the stream path instead of month:seconds — consistent with the Workbook path, which passes the pattern to Excel as a numFormat.
- **docs:** the docs-site package home pages (en + zh) still described the pre-2.0 dependency model (install `modern-xlsx` explicitly, peerDependencies incl. an optional `xlsx` fallback, mandatory asset deployment) — updated to the zero-runtime-dependency, zero-configuration reality; option tables document `workerTimeoutMs`; sheet-name constraints and the stream-route retry semantics are documented (en + zh).
- **tests:** 108 → 116 (empty-sheets pre-flight on both paths, postMessage-throw cleanup with no delayed worker kill, single stream retry on the stream route, diamond column reuse, apostrophe sheet names, custom worker timeout, `mm:ss` minute disambiguation).
