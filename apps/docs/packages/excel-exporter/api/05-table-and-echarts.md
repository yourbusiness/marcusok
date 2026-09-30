# API: Table & ECharts Adapters

Two convenience entry points that accept the data shapes you already have — a table `columns + data` pair, or a small structural subset of an ECharts option — normalize them into a [`SheetConfig`](/packages/excel-exporter/api/02-types#sheetconfig), and hand the result to [`exportExcel`](/packages/excel-exporter/api/01-export-excel). Both are thin adapters: everything they can express is expressible with `exportExcel` directly.

## Relationship with exportExcel

`exportTable` and `exportEcharts` are wrappers around `exportExcel`, in two steps:

1. a pure conversion (`tableExportToOptions` / `echartsExportToOptions`) turns the convenience options into an `ExportOptions`;
2. that result is passed to `exportExcel`.

Consequences worth knowing:

- The shared [`ExportOptions`](/packages/excel-exporter/api/01-export-excel#exportoptions) fields these two types do declare — `filename`, `mode`, `onProgress`, `onPhase`, `download` — behave exactly as they do on the main entry (same routing, same progress contract, same `{ success, blob, engine, mode, duration, rowCount, error }` result).
- The `overlay` field of `ExportOptions` is **not** part of `TableExportOptions` / `EChartsExportOptions`, and neither converter forwards it. The overlay therefore still runs with its default (on), but you cannot customize or disable it through these two wrappers — call `exportExcel` with a converted sheet if you need that.
- Both wrappers are single-sheet: each produces a `sheets` array of exactly one entry.

```ts
import {
  exportTable,
  exportEcharts,
  tableToSheet,
  tableExportToOptions,
  echartsToSheet,
  echartsExportToOptions,
} from "@marcusok/excel-exporter";
```

There are no subpath entries for these — the main entry re-exports the two adapters (`src/table-export.ts`, `src/echarts-export.ts`) in full, functions and types alike.

## exportTable

```ts
exportTable(options: TableExportOptions): Promise<ExportResult>
```

Exports one sheet from a `columns + data` pair, with column descriptors in the shapes Element Plus / Ant Design components already use. The sheet name defaults to `"Sheet1"`, and `freezeRows` / `autoFilter` / `merges` / `dataStyle` / `indexColumn` pass straight through to the sheet.

### TableExportOptions

`TableExportOptions` extends `TableSheetInput` with the export-level fields:

| Field         | Type                                               | Required | Description                                                                               |
| ------------- | -------------------------------------------------- | -------- | ----------------------------------------------------------------------------------------- |
| `columns`     | `TableColumnInput[]`                               | yes      | Column descriptors, normalized to `ColumnConfig` (see below)                              |
| `data`        | `Record<string, unknown>[]`                        | yes      | Row objects, keyed by the resolved `prop`                                                 |
| `filename`    | `string`                                           | yes      | Download name; `.xlsx` is appended unless already present                                 |
| `sheetName`   | `string`                                           | —        | Default `"Sheet1"`                                                                        |
| `freezeRows`  | `number`                                           | —        | Number of header rows to freeze; validated as a non-negative integer                      |
| `autoFilter`  | `boolean`                                          | —        | Header auto filter covering the last header row plus all data rows                        |
| `merges`      | `MergeRange[]`                                     | —        | Merged cell ranges, relative to the data area                                             |
| `dataStyle`   | `CellStyle`                                        | —        | Base style for every data cell; a column `style` deep-merges over it field by field       |
| `indexColumn` | `boolean \| IndexColumnOptions`                    | —        | Inject a leading row-number column; `true` for defaults. Merges shift right automatically |
| `mode`        | `"auto" \| "main" \| "worker" \| "stream"`         | —        | Default `"auto"`                                                                          |
| `onProgress`  | `(progress: number) => void`                       | —        | 0 → 1, with the leading 0 and trailing 1 each fired once on every route                   |
| `onPhase`     | `(phase: ExportPhase, durationMs: number) => void` | —        | `init` / `build` / `download` timings                                                     |
| `download`    | `boolean`                                          | —        | Default `true`; `false` returns the Blob only                                             |

### TableColumnInput

A deliberately dependency-free column descriptor — no Element Plus or Ant Design types leak into the signature.

| Field          | Type                     | Required | Description                                                                                         |
| -------------- | ------------------------ | -------- | --------------------------------------------------------------------------------------------------- |
| `prop?`        | `string`                 | leaf     | Field name on the data row (Element Plus / library naming); wins over `key` and `dataIndex`         |
| `key?`         | `string`                 | —        | Legacy alias of `prop` (pre-2.2 naming); used when `prop` is absent                                 |
| `dataIndex?`   | `string`                 | —        | Ant Design alias of `prop`; the last fallback                                                       |
| `label?`       | `string \| number`       | yes      | Header text; wins over `header` and `title`. A finite number is stringified                         |
| `header?`      | `string \| number`       | —        | Legacy alias of `label` (pre-2.2 naming)                                                            |
| `title?`       | `string \| number`       | —        | Ant Design alias of `label`; the last fallback                                                      |
| `width?`       | `number`                 | —        | Column width in Excel character units (`0` hides the column); leaf columns only                     |
| `style?`       | `CellStyle`              | —        | Data-cell style, deep-merged over the sheet `dataStyle`; leaf columns only                          |
| `headerStyle?` | `CellStyle`              | —        | Header style for this column (group header cells included); leaf and group columns alike            |
| `format?`      | `FormatSpec \| Function` | —        | Value formatting (see [FormatSpec](/packages/excel-exporter/api/03-format-spec)); leaf columns only |
| `children?`    | `TableColumnInput[]`     | —        | Group column: becomes a multi-row header and recurses. `children: []` counts as a leaf              |

Alias resolution is a plain `??` chain, so precedence is: `prop` → `key` → `dataIndex` for the field name, and `label` → `header` → `title` for the header text.

A column with a non-empty `children` array is a **group**: it produces a group header cell spanning its subtree and requires only a header text — `prop`, `width`, `style` and `format` are meaningless there (no data cells) and are dropped. Leaf headers span the remaining header rows vertically; group headers span horizontally. Header merges are generated automatically, exactly as described for `ColumnConfig` on the [Core Types](/packages/excel-exporter/api/02-types#columnconfig) page.

Validation happens during conversion, and fails fast with a specific message:

| Situation                                   | Error message                                                                                              |
| ------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| Leaf/group without a usable header text     | `[excel-exporter] table column "<prop or group-N>" has no usable header. Provide label, header, or title.` |
| Leaf with a header but no usable field name | `[excel-exporter] table column #<index> has no usable prop. Provide prop, key, or dataIndex.`              |
| `children` forming a cycle                  | `[excel-exporter] circular children reference in table columns`                                            |

Details that matter in practice:

- Header and prop are checked in that order, so a column with neither reports the header error first. In that message `<prop or group-N>` is the resolved prop when there is one, otherwise the placeholder `group-<index>` (`group-0` for the first top-level column).
- `<index>` is the column's position among its own siblings — a bad child reports its index inside `children`, not its absolute position.
- Empty strings do not qualify as headers: a header must be a non-empty string, or a finite number (which is stringified). `label: ""` throws.
- The cycle check exists because this conversion runs **before** `exportExcel` flattens the column tree: without it, a circular `children` reference would surface as a stack overflow instead of a readable error.

### TableSheetInput → SheetConfig

`tableToSheet` is the mapping, and it is nearly one-to-one:

| `TableSheetInput` | `SheetConfig` | Note                                             |
| ----------------- | ------------- | ------------------------------------------------ |
| `columns`         | `columns`     | Recursively normalized to `ColumnConfig`         |
| `data`            | `data`        | Passed through untouched                         |
| `sheetName`       | `name`        | Defaults to `"Sheet1"`                           |
| `freezeRows`      | `freezeRows`  | Forwarded when defined                           |
| `autoFilter`      | `autoFilter`  | Forwarded when defined                           |
| `merges`          | `merges`      | Forwarded when defined                           |
| `dataStyle`       | `dataStyle`   | Forwarded when defined                           |
| `indexColumn`     | `indexColumn` | Forwarded when defined; `exportExcel` expands it |

Only the five forwarding fields are conditional (`undefined` stays `undefined`, so the sheet keeps the engine default). When you need anything beyond this shape — multiple sheets, a sheet-level `headerStyle`, hand-written `ColumnConfig` trees — convert with `tableToSheet` and drive `exportExcel` (or [`WorkbookBuilder`](/packages/excel-exporter/api/01-export-excel#other-exported-symbols)) yourself.

### Complete example

```ts
import { exportTable, StylePresets } from "@marcusok/excel-exporter";

const result = await exportTable({
  filename: "quarterly-sales",
  sheetName: "Sales",
  freezeRows: 1,
  autoFilter: true,
  indexColumn: { label: "No.", width: 6 },
  dataStyle: StylePresets.dataRow,
  columns: [
    // Element Plus naming
    { prop: "orderId", label: "Order ID", width: 18 },
    // Ant Design naming
    { dataIndex: "date", title: "Date", width: 12, format: { type: "date" } },
    // Grouped columns: a multi-row header is generated for you
    {
      label: "Amount",
      headerStyle: StylePresets.header,
      children: [
        { prop: "net", label: "Net", width: 14, style: StylePresets.currency },
        { prop: "tax", label: "Tax", width: 12, style: StylePresets.currency },
      ],
    },
    {
      prop: "region",
      label: "Region",
      width: 12,
      format: {
        type: "enum",
        map: { apac: "APAC", emea: "EMEA", na: "North America" },
        fallback: "Other",
      },
    },
  ],
  data: [
    {
      orderId: "ORD-000001",
      date: "2026-07-01",
      net: 1099.99,
      tax: 199.99,
      region: "apac",
    },
    {
      orderId: "ORD-000002",
      date: "2026-07-02",
      net: 399,
      tax: 72,
      region: "emea",
    },
  ],
});

if (!result.success) console.error(result.error);
```

## exportEcharts

```ts
exportEcharts(options: EChartsExportOptions): Promise<ExportResult>
```

Exports one sheet from a small, explicit structural subset of an ECharts option. The adapter does not depend on the ECharts runtime or its type system, so you can pass the option object you already built — but only a documented subset is understood, and everything outside it is rejected rather than guessed.

### EChartsExportOptions

`EChartsExportOptions` extends `EChartsSheetInput` with the export-level fields:

| Field            | Type                                               | Required | Description                                                             |
| ---------------- | -------------------------------------------------- | -------- | ----------------------------------------------------------------------- |
| `option`         | `EChartsOptionInput`                               | yes      | The (subset of an) ECharts option to convert                            |
| `filename`       | `string`                                           | yes      | Download name; `.xlsx` is appended unless already present               |
| `sheetName`      | `string`                                           | —        | Default `"图表数据"`                                                    |
| `layout`         | `"wide" \| "long"`                                 | —        | Default `"wide"`; only meaningful for the category layout               |
| `categoryHeader` | `string`                                           | —        | Default `"类目"`                                                        |
| `seriesHeader`   | `string`                                           | —        | Default `"系列"`                                                        |
| `nameHeader`     | `string`                                           | —        | Default `"名称"`                                                        |
| `valueHeader`    | `string`                                           | —        | Default `"数值"`                                                        |
| `mode`           | `"auto" \| "main" \| "worker" \| "stream"`         | —        | Default `"auto"`                                                        |
| `onProgress`     | `(progress: number) => void`                       | —        | 0 → 1, with the leading 0 and trailing 1 each fired once on every route |
| `onPhase`        | `(phase: ExportPhase, durationMs: number) => void` | —        | `init` / `build` / `download` timings                                   |
| `download`       | `boolean`                                          | —        | Default `true`; `false` returns the Blob only                           |

### EChartsOptionInput and friends

| Type                 | Definition                                                                                                                                                 |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `EChartsScalar`      | `number \| string \| null`                                                                                                                                 |
| `EChartsDatum`       | `EChartsScalar \| number[] \| { name?: string; value?: EChartsScalar \| number[] }`                                                                        |
| `EChartsSeriesInput` | `{ name?: string; type?: string; data?: EChartsDatum[] }`                                                                                                  |
| `EChartsXAxisInput`  | `{ type?: string; data?: Array<number \| string> }`                                                                                                        |
| `EChartsOptionInput` | `{ xAxis?: EChartsXAxisInput \| EChartsXAxisInput[]; yAxis?: EChartsXAxisInput \| EChartsXAxisInput[]; series?: EChartsSeriesInput[]; dataset?: unknown }` |
| `EChartsLayout`      | `"wide" \| "long"`                                                                                                                                         |

Notes on the shapes:

- `type` on a series or axis is **ignored** — the layout is decided by the data, not by the declared chart type.
- `yAxis` shares `EChartsXAxisInput`: an axis is an axis structurally. It is consulted as the category source when `xAxis.data` is absent or empty, which is what makes horizontal bar charts work.
- `dataset` is declared only so the adapter can reject it explicitly.
- Series names default to `系列1`, `系列2`, … (1-based) when `name` is missing, empty, `null` or `undefined` — an empty string is treated as unnamed, so you never get an error that fails to point at a specific series.

### Supported data shapes

**Category layout** — one of the two axes carries `.data`, and each series has a one-dimensional `data` of the same length:

| Shape                                  | Example                                                                          | Result (default `wide`)                             |
| -------------------------------------- | -------------------------------------------------------------------------------- | --------------------------------------------------- |
| Vertical bar / line (categories on x)  | `xAxis: { data: ["Q1", "Q2"] }`, `series: [{ name: "Revenue", data: [10, 20] }]` | One row per category, one column per series         |
| Horizontal bar (categories on `yAxis`) | `yAxis: { data: [...] }`, no `xAxis.data`                                        | Same as above — `yAxis.data` is the category source |

The category source is `xAxis.data` when it is an array with at least one entry, otherwise `yAxis.data`. An empty `xAxis.data` (a chart still loading) counts as "not provided" and falls back to `yAxis.data` instead of silently dropping the categories.

**Item layout** — used when neither axis supplies categories. Which item layout you get is decided by the data, per series:

| Shape                | Example series data                                              | Columns produced                                                         |
| -------------------- | ---------------------------------------------------------------- | ------------------------------------------------------------------------ |
| Pie / name-value     | `[{ name: "Chrome", value: 62 }, { name: "Safari", value: 19 }]` | `系列` / `名称` / `数值`                                                 |
| Bare scalars         | `[62, 19]`                                                       | `系列` / `名称` / `数值` (name is the 1-based position, `"1"`, `"2"`, …) |
| Scatter pair         | `[[10, 20], [30, 40]]`                                           | `系列` / `X` / `Y`                                                       |
| Scatter, extra dims  | `[[10, 20, 5], [30, 40, 8]]`                                     | `系列` / `X` / `Y`, extra dims dropped with a one-time `console.warn`    |
| Scatter, object form | `[{ value: [10, 20] }, { value: [30, 40] }]`                     | `系列` / `X` / `Y`                                                       |

Details of the item layout:

- Scatter and name/value data cannot be mixed — see the rejection table below.
- The two scatter spellings (`[x, y]` and `{ value: [x, y] }`) are interchangeable and may be mixed with each other. An optional `name` on the object form is accepted and ignored: the scatter layout has no name column.
- A coordinate pair is any all-number array of length ≥ 2. Multi-dimensional scatter (`[x, y, ...dims]`, where dims drive `symbolSize` / `visualMap`) is exported as its first two dims, and the drop is reported **once per export** via `console.warn` rather than silently or per point:
  `[excel-exporter] scatter data has dimensions beyond [x, y]; only the X/Y coordinates are exported (extra dims drive symbolSize/visualMap and have no table column).`
- When only some series carry coordinates and others carry name/value data, the export is rejected (see below); the "is this scatter?" decision is driven by the presence of at least one real coordinate item, so a series with an empty `data` next to a scatter series does not turn a pie chart into a false scatter table.

### Rejected shapes

These are refused with an explicit error instead of producing a misleading table. All messages are prefixed with `[excel-exporter]`:

| Shape                                                      | Error message                                                                                                                                                                                                                                       |
| ---------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `option.dataset` present (dataset mode)                    | `ECharts dataset mode is not supported by echartsToSheet. Flatten the data before calling it.`                                                                                                                                                      |
| `option.series` missing, empty, or not an array of objects | `ECharts option has no series to export.` / `ECharts option.series must be an array of series objects (found a non-object element).`                                                                                                                |
| More than one `xAxis` or `yAxis` entry                     | `multiple x axes are not supported by echartsToSheet (a dual-axis chart has no single-columnar table shape).` (`y` for the y axis)                                                                                                                  |
| A series length differing from the category axis           | `ECharts series "系列1" must have the same length as the category axis data (xAxis.data / yAxis.data).`                                                                                                                                             |
| Object / coordinate-pair datum under a category axis       | `unsupported ECharts datum for category layout in series 0 (object/array data points are not supported with a category axis). Flatten each datum to a scalar, or drop xAxis.data/yAxis.data to export via the item (name/value or scatter) layout.` |
| Scatter and name/value data in the same option             | `mixing scatter coordinate data with name/value data is not supported by echartsToSheet.`                                                                                                                                                           |

The dual-axis rejection is deliberate: with two axes there is no single table shape, and silently taking the first axis would report the other axis's series as a baffling length mismatch.

Note that the category-layout datum rule still allows a _scalar_ datum, including a one-element array such as `[5]` (only an all-number array of length ≥ 2 counts as a coordinate pair). The reliable rule of thumb is: under a category axis, series data must be flat scalars.

### layout: wide vs long

`layout` is read **only** by the category layout. Item data (pie / scatter) ignores it entirely, because those layouts are already one row per datum.

| `layout`           | Shape                                                                              | Columns                                                       |
| ------------------ | ---------------------------------------------------------------------------------- | ------------------------------------------------------------- |
| `"wide"` (default) | One row per category; one column per series                                        | `类目` + one column per series, labelled with the series name |
| `"long"`           | One row per series-category pair (`{ seriesHeader, categoryHeader, valueHeader }`) | `系列`, `类目`, `数值`                                        |

Wide layout is what a spreadsheet reader expects for a small number of series; long layout is what pivots, filters and charts-with-many-series prefer.

### Headers

Default headers and the sheet name are Chinese, so they can be overridden for an English workbook via `sheetName` / `seriesHeader` / `categoryHeader` / `nameHeader` / `valueHeader`. Two constraints come with that:

- **Long and item layouts require distinct header texts**, because the header strings double as row keys (`{ [seriesHeader]: name, [categoryHeader]: category, [valueHeader]: value }`). A duplicate would silently overwrite a column, so it is rejected up front — e.g. `duplicate header "系列" in category long layout: header texts double as row keys in long/item layouts, so they must be distinct (rename via the *Header options).` (the context string is `name/value layout` or `scatter layout` for the item layouts).
- **Wide layout rejects a `categoryHeader` that collides with the internal series keys**: the wide rows are keyed by `categoryHeader` plus `__series_0`, `__series_1`, …, so a `categoryHeader` literally equal to one of those names would be overwritten by series data. That is rejected with `categoryHeader "__series_0" collides with the internal series keys (__series_N) in wide layout; choose a different categoryHeader.`

The scatter layout's coordinate headers are the literal `X` / `Y` and are not configurable; only `seriesHeader` applies to it.

### Complete example

```ts
import { exportEcharts } from "@marcusok/excel-exporter";

// Whatever chart you already render — only the documented subset is read.
const option = {
  xAxis: { type: "category", data: ["Q1", "Q2", "Q3", "Q4"] },
  series: [
    { name: "Revenue", type: "bar", data: [120, 200, 150, 80] },
    { name: "Cost", type: "bar", data: [90, 140, 120, 60] },
  ],
};

const result = await exportEcharts({
  filename: "quarterly-chart",
  option,
  sheetName: "Chart data",
  layout: "long", // one row per series-category pair
  seriesHeader: "Series",
  categoryHeader: "Quarter",
  valueHeader: "Amount",
});

if (!result.success) console.error(result.error);
```

The same option with the defaults (`layout: "wide"`, Chinese headers) exports a sheet named `图表数据` with the columns `类目` / `Revenue` / `Cost`.

## Low-level converters

All four functions are exported from the main entry. They are pure — no download, no overlay, no `exportExcel` import behind them — which makes them the seam to use when you want to inspect, assert, or post-process the generated configuration before driving a lower-level entry point yourself.

| Function                                              | Input                  | Output          | Use it when                                                                                                                  |
| ----------------------------------------------------- | ---------------------- | --------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `tableToSheet(input: TableSheetInput)`                | `TableSheetInput`      | `SheetConfig`   | You want the table shape but a different workbook: combine several sheets, or feed `WorkbookBuilder` / `exportAsStream`      |
| `tableExportToOptions(input: TableExportOptions)`     | `TableExportOptions`   | `ExportOptions` | You want to tweak the resulting `ExportOptions` (e.g. add a second sheet or an `overlay` value) before calling `exportExcel` |
| `echartsToSheet(input: EChartsSheetInput)`            | `EChartsSheetInput`    | `SheetConfig`   | Same as `tableToSheet`, for chart data                                                                                       |
| `echartsExportToOptions(input: EChartsExportOptions)` | `EChartsExportOptions` | `ExportOptions` | Same as `tableExportToOptions`, for chart data                                                                               |

Two caveats when you drive the lower-level entry points yourself:

- `WorkbookBuilder.addSheet()` and `exportAsStream()` do **not** expand `indexColumn`; pass `applyIndexColumn(sheet)` (also exported) if the converted sheet uses it — see [IndexColumnOptions](/packages/excel-exporter/api/02-types#indexcolumnoptions).
- The converters transfer the same fields the wrappers do — including the fact that `overlay` is not among them, so the sheet you get back is plain configuration.

```ts
import { tableToSheet, exportExcel } from "@marcusok/excel-exporter";

const sheet = tableToSheet({
  sheetName: "Sales",
  columns: [{ prop: "region", label: "Region" }],
  data: [{ region: "APAC" }],
});

// Inspect or adjust the generated config, then export it alongside others.
await exportExcel({
  filename: "combined",
  sheets: [
    sheet,
    { name: "Notes", columns: [{ prop: "body", label: "Note" }], data: [] },
  ],
  overlay: false, // only reachable because this is the main entry
});
```

## Error contract

The two wrappers never reject on a conversion error. Their option types promise `Promise<ExportResult>`, and a structural problem in the input — an unnamed column, a `dataset` option, a series length mismatch, a duplicate header — is returned as a normal failed result:

```ts
const result = await exportEcharts({ filename: "x", option: { dataset: {} } });

result.success; // false
result.error; // Error: [excel-exporter] ECharts dataset mode is not supported ...
```

This matches `exportExcel`, which already reports its own failures as `{ success: false, error }`: the conversion errors simply happen _before_ `exportExcel` is reached, so the wrappers catch them and wrap them in the same shape. The same bad input therefore fails identically whether it arrives through a convenience wrapper or through a hand-built `ExportOptions`.

Note that `success: false` also covers the fallback path — check `result.error` (and `result.engine` / `result.mode`) rather than assuming a particular engine was used. See [ExportResult](/packages/excel-exporter/api/01-export-excel#exportresult).
