---
"@marcusok/excel-preview": patch
---

Speed up the per-scroll-frame text-spill scan on sparse wide rows: `spillLimitPx` now advances through the row's sorted cell array after a single binary search instead of binary-searching every gap column, and a new performance baseline guards the sparse-string worst case (60 rows x 16384 cols).
