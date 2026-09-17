# Styling

A column-level `style` (`CellStyle`) applies to all **data cells** of that column — headers are styled separately via `headerStyle` (column-level) or the sheet-level `headerStyle`, and the sheet-level `dataStyle` sets a base layer for every data cell (see below). Eight presets are built in, and full customization is supported.

## Built-in presets

| Preset                  | Visual                                                                                                                   | Description                                                                                    |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------- |
| `StylePresets.header`   | <span style="display:inline-block;width:12px;height:12px;background:#1F4E79;border-radius:2px"></span> dark blue / white | Bold, size 12, `1F4E79` fill, white text, centered                                             |
| `StylePresets.currency` | `#,##0.00`                                                                                                               | Thousands separator, 2 decimals, right-aligned                                                 |
| `StylePresets.percent`  | `0.00%`                                                                                                                  | Percentage format, right-aligned                                                               |
| `StylePresets.date`     | `yyyy-MM-dd`                                                                                                             | Date format, centered                                                                          |
| `StylePresets.datetime` | `yyyy-MM-dd HH:mm`                                                                                                       | Date-time format, centered                                                                     |
| `StylePresets.dataRow`  | left + thin bottom border                                                                                                | Left-aligned, vertically centered, thin `D0D0D0` bottom border                                 |
| `StylePresets.bordered` | thin box on all four sides                                                                                               | Thin `D0D0D0` borders on all sides — pairs with sheet-level `dataStyle` for table-wide borders |
| `StylePresets.danger`   | <span style="display:inline-block;width:12px;height:12px;background:#C00000;border-radius:2px"></span> red bold          | Bold red text `C00000`, centered                                                               |

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "styled",
  sheets: [
    {
      name: "Sheet1",
      columns: [
        { prop: "name", label: "Name", width: 16, style: StylePresets.dataRow },
        {
          prop: "amount",
          label: "Amount",
          width: 14,
          style: StylePresets.currency,
        },
        { prop: "date", label: "Date", width: 12, style: StylePresets.date },
        {
          prop: "status",
          label: "Status",
          width: 10,
          style: StylePresets.danger,
        },
      ],
      data: [
        {
          name: "Keyboard",
          amount: 1299.99,
          date: "2026-07-01",
          status: "Out of stock",
        },
      ],
    },
  ],
});
```

## Custom CellStyle

```ts
import type { CellStyle } from "@marcusok/excel-exporter";

const highlight: CellStyle = {
  font: { bold: true, size: 11, color: "1F4E79" }, // 6-digit RGB hex
  fill: { pattern: "solid", fgColor: "DDEBF7" },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  border: {
    bottom: { style: "medium", color: "1F4E79" },
    right: { style: "thin", color: "D0D0D0" },
  },
  numFormat: "#,##0.00",
};
```

Field reference:

| Field       | Description                                                                                          |
| ----------- | ---------------------------------------------------------------------------------------------------- |
| `font`      | `bold` / `italic` / `size` / `color` (6-digit hex, e.g. `"FF0000"`) / `name`                         |
| `fill`      | `pattern: "solid" \| "none"`, `fgColor`, `bgColor` (6-digit hex)                                     |
| `alignment` | `horizontal` (left/center/right), `vertical` (top/center/bottom), `wrapText`, `textRotation` (0–180) |
| `border`    | per-side `{ style, color }`; `style` values come from modern-xlsx `BorderStyle`                      |
| `numFormat` | Excel format codes such as `"#,##0.00"`, `"yyyy-mm-dd"`, `"0.00%"`                                   |

> Colors are 6-digit RGB hex **without** `#`, matching modern-xlsx's type contract.

## Sheet-level dataStyle

`SheetConfig.dataStyle` is the base style for **every data cell** — one field covers the whole table, no per-column repetition:

```ts
await exportExcel({
  filename: "bordered",
  sheets: [
    {
      name: "Sheet1",
      headerStyle: StylePresets.header, // headers
      dataStyle: StylePresets.bordered, // all data cells, one field
      columns: [
        { prop: "name", label: "Name" },
        // A column style deep-merges over dataStyle: the table-wide border
        // survives, the numFormat / alignment come from the column.
        { prop: "amount", label: "Amount", style: StylePresets.currency },
      ],
      data,
    },
  ],
});
```

The merge is **field-level**, not wholesale: `dataStyle` provides the base layer and a column's `style` overrides only the fields it sets — a table-wide border survives a column that only sets `numFormat`, and vice versa. This is deliberately different from `headerStyle`, which replaces wholesale. Like all styling, `dataStyle` is dropped with a warning on the stream path (≥ 50,000 rows / degraded exports).

## Index column

`SheetConfig.indexColumn` injects a leading row-number column — no data preprocessing, no extra entry in `columns`:

```ts
sheets: [
  {
    name: "Sheet1",
    indexColumn: true, // or { label: "No.", width: 6, start: 1, style, headerStyle }
    columns: [
      { prop: "name", label: "Name" },
      { prop: "amount", label: "Amount" },
    ],
    data, // untouched — numbers are generated from the row number
  },
];
```

- Works on **every export path** (workbook / worker / stream): it is structure, not styling, so even the style-less stream keeps the numbers.
- Existing `merges` are shifted one column right automatically, so they keep pointing at their original targets.
- Values never read `data`; a user column declaring the reserved `__index__` prop is rejected with a clear error.
- Styling follows the same rules as any column: header via `indexColumn.headerStyle` (over sheet `headerStyle`), data cells via `indexColumn.style` (merged over sheet `dataStyle`).

See the [API reference](../api/02-types) for the full option list.

## Relationship to FormatSpec

`style.numFormat` and column `format` (FormatSpec) are separate mechanisms: `style` controls appearance, `format` converts/structures the stored value. For `date` / `number` FormatSpecs the Workbook path auto-injects the matching `numFormat`, so manual setup is usually unnecessary. See [Formatting](/packages/excel-exporter/guide/04-formatting).
