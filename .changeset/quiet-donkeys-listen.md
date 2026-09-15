---
"@marcusok/excel-exporter": patch
---

Docs accuracy pass (no behavior change):

- README: `onPhase` now documented as per-attempt (a degradation chain reports one `build` phase per attempt, matching the `ExportPhase` type contract); the WASM-needs table entry now includes auto/Node >= 50,000 rows alongside explicit `mode: "stream"` as WASM-free routes.
- `types.ts`: the `FormatSpec` / `ColumnConfig.format` comments now state precisely where function-form formats run (all main-thread routes, including Node's main-thread stream and degradation retries) and where they are stripped with a warning (browser worker routes).
