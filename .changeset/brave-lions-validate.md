---
"@marcusok/excel-exporter": patch
---

- Input validation: a misspelled `format.type` (e.g. `"percent"`) now fails fast with a clear error instead of silently falling through `applyFormat`'s default branch, which stringified the whole column into text cells with `success: true`.
- Input validation: `format.map` / `format.pattern` / `format.fill` shape errors and a non-array `merges` value are now rejected up front with column/sheet-locating messages, instead of surfacing as raw TypeErrors after a wasted style-less stream fallback.
- Cross-path parity: empty-string cells (normalized missing values) are now written as physical blanks (no `<c>` element) on the Workbook path, matching the stream path — `ISBLANK()`/`COUNTA` no longer flip when a dataset crosses the 50k-row threshold.
- ECharts adapter: sparse `option.series` arrays are rejected with a clear error (previously a raw TypeError in the item layout); an empty `xAxis.data` array now falls back to `yAxis.data` instead of silently switching to the item layout; empty-string series names fall back to the default `系列N` instead of failing column validation without a series reference.
- `exportTable` / `exportEcharts`: conversion-time input errors now resolve as `{ success: false, error }` like `exportExcel`, instead of rejecting the promise — the same bad input fails the same way through every entry point.
