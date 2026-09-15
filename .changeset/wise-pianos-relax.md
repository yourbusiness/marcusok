---
"@marcusok/excel-exporter": patch
---

Docs and filename edge case:

- `download.ts`: the `.xlsx` suffix check is now case-insensitive, so a filename like `report.XLSX` no longer becomes `report.XLSX.xlsx`.
- `types.ts` doc comments: `FormatSpec` now documents that the stream path parses only the `yyyy`/`MM`/`dd`/`HH`/`mm`/`ss` pattern tokens and emits anything else verbatim (the Workbook path renders any valid Excel format code via numFormat); `ExportPhase` now states that the browser worker-stream route reports no `init` phase and that a build attempt failing inside the worker reports no `build` phase (main-thread attempts report in a `finally`); the `onPhase` note on `ExportResult.duration` now reflects that the worker route measures in-worker time only.
