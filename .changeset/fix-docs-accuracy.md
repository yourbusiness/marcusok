---
"@marcusok/excel-exporter": patch
---

Docs: fix inaccurate statements and document missing behavior details.

- The ECharts category-layout rule now states that any array is rejected, including the one-element `[5]` spelling (previously documented as allowed).
- The fallback note now matches the actual contract: a degraded-but-successful export is `success: true` with `error` set (previously implied `success: false`).
- The "split into multiple sheets" styling advice is corrected: `auto` routes on the per-call total across sheets; splitting means multiple `exportExcel` calls.
- The `configureWasm` example no longer shows a CDN `workerUrl` (worker scripts are same-origin only; the same note is added to the option tables).
- Newly documented: enum missing values render as empty cells and never reach `map`/`fallback` (also fixes the stale `ColumnConfig.format` comment in `src/types.ts` to match the implementation and tests); up-front validation rules for `CellStyle` fields, `MergeRange`, case-insensitive sheet-name uniqueness and the `mode` enum; `ExportResult.duration` and `ExportPhase` per-route reporting semantics; the `toStr` export; superset-token / unparseable-value edges of date specs; `workerTimeoutMs` having no disable value (`0` means immediate timeout); and the URL-changes-take-effect-only-in-a-fresh-JS-realm caveat on the installation guide.
