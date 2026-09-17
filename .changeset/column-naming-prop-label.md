---
"@marcusok/excel-exporter": minor
---

Align column field naming with Element Plus: `prop` (data-row field) and `label` (header text) are now the canonical `ColumnConfig` fields. The pre-2.2 names `key` / `header` keep working as deprecated aliases — `prop` / `label` win when both are present, so existing code exports unchanged. `exportTable()` resolves names as `prop ?? key ?? dataIndex` and `label ?? header ?? title`. Columns providing neither `label` nor `header` are now rejected up front with a clear error.
