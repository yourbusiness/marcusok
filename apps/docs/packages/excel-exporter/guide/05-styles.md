# Styling

A column-level `style` (`CellStyle`) applies to all **data cells** of that column — headers are styled separately via `headerStyle` (column-level) or the sheet-level `headerStyle`, and the sheet-level `dataStyle` sets a base layer for every data cell (see below). Eight presets are built in, and full customization is supported — including deriving your own variants from the presets with object spread.

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

## Putting it together

The presets are designed to combine: one sheet-level `headerStyle`, one sheet-level `dataStyle`, and a couple of column-level styles produce a complete report look:

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "monthly-report",
  sheets: [
    {
      name: "Orders",
      headerStyle: StylePresets.header, // dark-blue headers, whole table
      dataStyle: StylePresets.bordered, // thin borders on every data cell
      indexColumn: { label: "No.", width: 6 }, // leading row-number column
      freezeRows: 1,
      autoFilter: true,
      columns: [
        { prop: "orderDate", label: "Date", style: StylePresets.date },
        { prop: "amount", label: "Amount", style: StylePresets.currency },
        { prop: "rate", label: "Growth", style: StylePresets.percent },
        { prop: "status", label: "Status", style: StylePresets.danger },
      ],
      data,
    },
  ],
});
```

Where each visual element comes from:

| Visual element                                   | Field                    |
| ------------------------------------------------ | ------------------------ |
| Dark-blue header, whole table                    | sheet `headerStyle`      |
| Thin border on every data cell                   | sheet `dataStyle`        |
| Leading `No.` column                             | sheet `indexColumn`      |
| `yyyy-MM-dd` dates, `#,##0.00` amounts, percents | the column's own `style` |

The column styles merge **over** `dataStyle` field by field (see [below](#sheet-level-datastyle)): `StylePresets.currency` only sets `numFormat` + `alignment`, so its cells keep the table-wide border from `dataStyle` and gain the number format from the column.

## Customizing a preset

Presets are plain constant objects (`as const`, typed via `satisfies CellStyle`) — not factories, not frozen. Derive your own variant with object spread instead of rewriting a `CellStyle` from scratch:

```ts
// Amounts with thousands separator but no decimals
style: { ...StylePresets.currency, numFormat: "#,##0" }

// Bold amounts (currency sets no font, so there is nothing to lose)
style: { ...StylePresets.currency, font: { bold: true } }

// Same header preset in a different brand color
headerStyle: { ...StylePresets.header, fill: { pattern: "solid", fgColor: "2E7D32" } }

// Index column: horizontally centered numbers, keeping dataRow's vertical
// centering (the nested alignment is spread — see the shallow-merge rule below)
indexColumn: {
  style: {
    ...StylePresets.dataRow,
    alignment: { ...StylePresets.dataRow.alignment, horizontal: "center" },
  },
}
```

::: warning Spread is a shallow merge
Object spread replaces **top-level fields wholesale**. Nested objects (`font`, `fill`, `alignment`, `border`) must be spread themselves, or their siblings are lost:

```ts
// ✗ loses bold / size / color — font is replaced entirely
{ ...StylePresets.header, font: { size: 14 } }

// ✓ only size changes, the rest of the font survives
{ ...StylePresets.header, font: { ...StylePresets.header.font, size: 14 } }
```

This is your own merge and happens before the library sees anything — different from the engine's field-level **deep** merge between `dataStyle` and a column's `style` (see below).
:::

Common `numFormat` tweaks used with the numeric presets:

| Format code          | Renders                |
| -------------------- | ---------------------- |
| `#,##0`              | `12,999` (no decimals) |
| `0%`                 | `42%` (no decimals)    |
| `"¥"#,##0.00`        | `¥12,999.99`           |
| `yyyy"年"M"月"d"日"` | `2026年7月1日`         |

Literal text in a format code goes in double quotes; everything else follows Excel's format-code syntax.

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

## Header styling: sheet default, column override

Headers have the same two levels, but with the opposite merge semantics — **wholesale replacement**, not field-level merging:

```ts
sheets: [
  {
    name: "Sheet1",
    headerStyle: StylePresets.header, // default for every header cell
    columns: [
      { prop: "name", label: "Name" },
      // A column-level headerStyle replaces the sheet-level one entirely
      // for this column's header cell — it does not merge with it.
      {
        prop: "amount",
        label: "Amount",
        headerStyle: {
          fill: { pattern: "solid", fgColor: "DDEBF7" },
          font: { bold: true, color: "1F4E79" },
        },
      },
    ],
    data,
  },
],
```

- Group columns (with `children`) accept `headerStyle` too — it styles that group's merged header cell.
- To derive from the default instead of rewriting it, spread it: `headerStyle: { ...StylePresets.header, fill: { pattern: "solid", fgColor: "DDEBF7" } }` — see [Customizing a preset](#customizing-a-preset) for the shallow-merge caveat.

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

## Where styles apply (and where they don't)

- **Styles render on the Workbook routes only**: auto mode under 50,000 rows (main thread, or Worker + Workbook for 20,000–49,999 rows in a browser). The stream path (≥ 50,000 rows, or a degraded export) strips all styling — `style` / `headerStyle` / `dataStyle` / `width` / `freezeRows` — with a console warning. Do not rely on styles for large-file exports.
- **Colors are 6-digit RGB hex without `#`** (`"1F4E79"`), matching modern-xlsx's type contract.
- **Reusing a preset is free**: the engine deduplicates structurally identical styles into a single style index, so referencing `StylePresets.currency` from N columns does not bloat the file — no need to share objects manually.

## Relationship to FormatSpec

`style.numFormat` and column `format` (FormatSpec) are separate mechanisms: `style` controls appearance, `format` converts/structures the stored value. For `date` / `number` FormatSpecs the Workbook path auto-injects the matching `numFormat`, so manual setup is usually unnecessary. See [Formatting](/packages/excel-exporter/guide/04-formatting).
