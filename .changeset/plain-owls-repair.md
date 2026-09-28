---
"@marcusok/excel-exporter": patch
---

- ECharts adapter: series with empty or missing `data` are now classified as the name/value layout instead of being misdetected as scatter — `every()` is vacuously true on an empty array, so a not-yet-loaded pie/bar series exported with `系列/X/Y` headers instead of `系列/名称/数值`.
- ECharts adapter: the category-layout error for object/array data points no longer suggests switching to the long layout (the check runs before the layout branch, so that advice never worked); it now suggests flattening each datum to a scalar or dropping the category axis to use the item layout.
- Duplicate sheet names are now rejected case-insensitively on every route (`"Sheet1"`/`"SHEET1"` previously passed validation), matching Excel's own case-insensitive uniqueness rule.
- Overlay: a render exception thrown during overlay setup can no longer strand the shared overlay's reference count, which would have left a revealed overlay on screen (and blocking interaction) forever.
