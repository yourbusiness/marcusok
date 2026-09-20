---
"@marcusok/excel-exporter": patch
---

Fix silent data loss when a user column claims the reserved `__index__` prop without `indexColumn` being enabled.

The builders treat `__index__` (or the legacy `key: "__index__"` alias) as the virtual row-number column **unconditionally** — such a column's cells are generated from the row number and never read from `data`. The conflict check, however, only ran inside `applyIndexColumn`, which returns the sheet untouched when `indexColumn` is not set. As a result, a hand-written `__index__` column passed to `exportExcel` was silently replaced by row numbers: the export reported `success: true` and no warning, and the user's values were gone.

The check now runs in the entry's pre-flight validation, so `exportExcel` rejects such a column with a structured error (`success: false`, `column prop "__index__" is reserved for the index column; …`) on every route — `main`, `worker` and `stream` alike — whether or not `indexColumn` is enabled. This matches what the package README and the docs site already documented.

The lower-level entry points (`WorkbookBuilder.addSheet`, `exportAsStream`) are unchanged: they perform no input validation and keep their documented "a hand-written `__index__` column is driven by the row number" semantics. Declaration stays the way to get an index column on those paths (`indexColumn` + `applyIndexColumn`).
