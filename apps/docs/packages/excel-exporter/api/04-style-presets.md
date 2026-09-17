# API: StylePresets

## Presets

| Name       | numFormat          | Other styles                                                   | Use for                         |
| ---------- | ------------------ | -------------------------------------------------------------- | ------------------------------- |
| `header`   | —                  | Bold, size 12, `1F4E79` fill, white text, centered             | Headers (when applied manually) |
| `currency` | `#,##0.00`         | Right-aligned                                                  | Amounts                         |
| `percent`  | `0.00%`            | Right-aligned                                                  | Ratios, growth rates            |
| `date`     | `yyyy-MM-dd`       | Centered                                                       | Date columns                    |
| `datetime` | `yyyy-MM-dd HH:mm` | Centered                                                       | Date-time columns               |
| `dataRow`  | —                  | Left-aligned, vertically centered, thin `D0D0D0` bottom border | Data rows                       |
| `bordered` | —                  | Thin `D0D0D0` box on all four sides                            | Table-wide borders              |
| `danger`   | —                  | Bold red `C00000`, centered                                    | Risk / anomalies                |

## Full definitions

The presets are plain constant objects — the exact values shipped in `src/style-presets.ts`:

```ts
export const StylePresets = {
  header: {
    font: { bold: true, size: 12, color: "FFFFFF" },
    fill: { pattern: "solid", fgColor: "1F4E79" },
    alignment: { horizontal: "center", vertical: "center" },
  },
  currency: { numFormat: "#,##0.00", alignment: { horizontal: "right" } },
  percent: { numFormat: "0.00%", alignment: { horizontal: "right" } },
  date: { numFormat: "yyyy-MM-dd", alignment: { horizontal: "center" } },
  datetime: {
    numFormat: "yyyy-MM-dd HH:mm",
    alignment: { horizontal: "center" },
  },
  dataRow: {
    alignment: { horizontal: "left", vertical: "center" },
    border: { bottom: { style: "thin", color: "D0D0D0" } },
  },
  bordered: {
    border: {
      top: { style: "thin", color: "D0D0D0" },
      bottom: { style: "thin", color: "D0D0D0" },
      left: { style: "thin", color: "D0D0D0" },
      right: { style: "thin", color: "D0D0D0" },
    },
  },
  danger: {
    font: { color: "C00000", bold: true },
    alignment: { horizontal: "center" },
  },
} as const;
```

## Usage

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

columns: [
  { prop: "amount", label: "Amount", width: 14, style: StylePresets.currency },
  { prop: "rate", label: "Growth", width: 12, style: StylePresets.percent },
  { prop: "date", label: "Date", width: 12, style: StylePresets.date },
  { prop: "flag", label: "Status", width: 10, style: StylePresets.danger },
];
```

## Where to attach a style

| Target                                       | Field                                       | Merge semantics                              |
| -------------------------------------------- | ------------------------------------------- | -------------------------------------------- |
| One column's data cells                      | `ColumnConfig.style`                        | deep-merges over sheet `dataStyle`           |
| One column's header (group headers included) | `ColumnConfig.headerStyle`                  | replaces sheet `headerStyle` wholesale       |
| Every data cell (base layer)                 | `SheetConfig.dataStyle`                     | overridable field by field by column `style` |
| Every header cell (default)                  | `SheetConfig.headerStyle`                   | replaced wholesale by column `headerStyle`   |
| The injected index column                    | `IndexColumnOptions.style` / `.headerStyle` | same rules as a regular column               |

## Deriving variants

Presets are not factories and not frozen — derive variants with object spread:

```ts
// Amounts without decimals
style: { ...StylePresets.currency, numFormat: "#,##0" }

// Only change one nested field: spread the nested object too,
// otherwise its siblings are lost (spread is a shallow merge)
style: { ...StylePresets.header, font: { ...StylePresets.header.font, size: 14 } }
```

The spread is your merge, done before the library sees the config — distinct from the engine's field-level deep merge between `dataStyle` and a column's `style`. Ready-made recipes (report template, header overrides, `numFormat` tweaks): see the [Styling guide](/packages/excel-exporter/guide/05-styles).

## Type

```ts
import type { StylePresetName } from "@marcusok/excel-exporter";

const name: StylePresetName = "currency"; // "header" | "currency" | "percent" | "date" | "datetime" | "dataRow" | "bordered" | "danger"
```

`StylePresetName` parameterizes preset selection:

```ts
const columnPresets: Record<string, StylePresetName> = {
  amount: "currency",
  rate: "percent",
  orderDate: "date",
};

const columns = Object.entries(columnPresets).map(([prop, name]) => ({
  prop,
  label: prop,
  style: StylePresets[name],
}));
```

> Column `style` applies to data cells, not headers. For header styling use the `headerStyle` field directly (sheet-level `SheetConfig.headerStyle` sets the default; column-level `ColumnConfig.headerStyle` overrides it), e.g. `headerStyle: StylePresets.header`.
