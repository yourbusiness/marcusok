---
"@marcusok/excel-exporter": patch
---

- fix: reject a sheet with an empty `columns` array up front with a clear error (`at least one column`) on every path, instead of crashing the Workbook autoFilter layout with a cryptic `TypeError` (`encodeCellRef(0, -1)` → `"@1"`) and then silently degrading to the SheetJS fallback.
- fix: `echartsToSheet` long/item layouts now reject duplicated header texts (e.g. `seriesHeader` equal to `valueHeader`) with a clear error — header texts double as row keys there, so duplicates previously overwrote each other's column silently.
- docs: correct the stale test-count figure in the README and document the Chinese default headers of `exportEcharts` (overridable via the `*Header` options).
