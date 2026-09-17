---
"@marcusok/excel-exporter": minor
---

Refine the default look of styled exports:

- **Default alignment — every cell is centered.** A base style (`BaseCellStyle`, horizontal + vertical `center`; also exported from the package) is layered underneath all cells on the Workbook path, headers included, so a table no longer mixes Excel's native alignment (text left, numbers right) column by column. Explicit `alignment` values still win field by field — declaring only `horizontal` inherits the base's `vertical` — and either `dataStyle` or `headerStyle` overrides the base. To restore Excel's native alignment, declare both axes explicitly (e.g. `dataStyle: { alignment: { horizontal: "left", vertical: "bottom" } }`).
- **Index column header text.** `SheetConfig.indexColumn` now defaults its `label` to `"序号"` instead of `"No."`. Pass an explicit `label` to display something else.
- **Darker preset borders.** `StylePresets.dataRow` and `StylePresets.bordered` now use `BFBFBF` instead of `D0D0D0`, so their thin borders read clearly rather than washing out.
