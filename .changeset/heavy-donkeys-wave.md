---
"@marcusok/excel-exporter": minor
---

Add sheet-level table styling and a built-in index column:

- `SheetConfig.dataStyle`: base `CellStyle` applied to every data cell. A column's `style` deep-merges over it field by field — a table-wide border survives a column that only sets `numFormat`, and vice versa. Together with the new `StylePresets.bordered` preset (thin `D0D0D0` box on all four sides), a fully bordered table is now `dataStyle: StylePresets.bordered`: one field, no per-column repetition.
- `SheetConfig.indexColumn` (`true`, or `{ label, width, start, style, headerStyle }`): injects a leading row-number column. Values are generated from the row number (never read from `data`), existing `merges` shift one column right automatically, and the feature works on every export path — workbook, worker, and the style-less stream alike. `exportTable` accepts both new fields too.
