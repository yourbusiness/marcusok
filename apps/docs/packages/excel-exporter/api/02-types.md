# API: Core Types

## SheetConfig

| Field          | Type                            | Required | Description                                                                                              |
| -------------- | ------------------------------- | -------- | -------------------------------------------------------------------------------------------------------- |
| `name`         | `string`                        | yes      | Non-empty, ≤ 31 chars, no `: \ / ? * [ ]`, no leading/trailing apostrophe                                |
| `columns`      | `ColumnConfig[]`                | yes      | Column definitions                                                                                       |
| `data`         | `Record<string, unknown>[]`     | yes      | Row data                                                                                                 |
| `headerStyle?` | `CellStyle`                     | —        | Default header style; overridden by column-level `headerStyle`                                           |
| `dataStyle?`   | `CellStyle`                     | —        | Base style for every data cell; column `style` deep-merges over it field by field (see the Styles guide) |
| `indexColumn?` | `boolean \| IndexColumnOptions` | —        | Inject a leading row-number column; `true` = all defaults. Merges shift right automatically              |
| `freezeRows?`  | `number`                        | —        | Freeze the first N header rows; validated as a non-negative integer                                      |
| `merges?`      | `MergeRange[]`                  | —        | Merged cells (relative to the data area)                                                                 |
| `autoFilter?`  | `boolean`                       | —        | Header auto filter                                                                                       |

## ColumnConfig

| Field          | Type                     | Required     | Description                                                                                                                                   |
| -------------- | ------------------------ | ------------ | --------------------------------------------------------------------------------------------------------------------------------------------- |
| `prop?`        | `string`                 | leaf columns | Field name on the data row (Element Plus naming); group columns (with `children`) may omit it                                                 |
| `key?`         | `string`                 | —            | Deprecated alias of `prop` (pre-2.2 naming); `prop` wins when both are present                                                                |
| `label?`       | `string`                 | yes*         | Header text (leaf and group columns alike); `label` or legacy `header` — at least one required (*validated at export time)                    |
| `header?`      | `string`                 | —            | Deprecated alias of `label` (pre-2.2 naming); `label` wins when both are present                                                              |
| `children?`    | `ColumnConfig[]`         | —            | Group column: produces a multi-row header; its header cell merges across all descendant leaf columns. `children: []` is a leaf                |
| `width?`       | `number`                 | —            | Column width (Excel character units; `0` hides the column); validated as a finite non-negative number; leaf columns only                      |
| `style?`       | `CellStyle`              | —            | Data-cell style (headers excluded); leaf columns only                                                                                         |
| `headerStyle?` | `CellStyle`              | —            | Header style for this column (group header cells included); wins over sheet-level `headerStyle`                                               |
| `format?`      | `FormatSpec \| Function` | —            | Value formatting; leaf columns only; functions run on main-thread paths and are stripped on the browser worker path (see the FormatSpec page) |

A column with `children` is a group: no data cells, header rows only. Header row count = 1 + the deepest column's tree depth (so 1 for a flat column list); leaf headers span the remaining header rows vertically, group headers span their leaf subtree horizontally — merges are generated automatically (no manual `merges` needed for headers).

## IndexColumnOptions

Options for `SheetConfig.indexColumn`; the shorthand `true` equals `{}`.

| Field          | Type        | Default  | Description                                                                                      |
| -------------- | ----------- | -------- | ------------------------------------------------------------------------------------------------ |
| `label?`       | `string`    | `"序号"` | Header text of the index column                                                                  |
| `width?`       | `number`    | `6`      | Column width (Excel character units; `0` hides the column)                                       |
| `start?`       | `number`    | `1`      | Number shown on the first data row; row i displays `start + i`. Non-negative integer (validated) |
| `style?`       | `CellStyle` | —        | Data-cell style of the index column; merged over sheet-level `dataStyle` like any column style   |
| `headerStyle?` | `CellStyle` | —        | Header style of the index column; overrides sheet-level `headerStyle`                            |

The index column's values are generated from the row number and never read from `data`; a user column declaring the reserved `__index__` prop is rejected with a clear error. Existing `merges` are shifted one column right so they keep pointing at their original targets.

`exportExcel` expands this field automatically (once, before mode routing), which is what makes it work on every route. The lower-level entry points — `WorkbookBuilder.addSheet()` and `exportAsStream()` — do **not** expand it: they only understand an already-expanded `__index__` column, so pass `applyIndexColumn(sheet)` (exported from the package, along with the reserved `INDEX_PROP`) when calling them directly. Otherwise `indexColumn` is silently ignored, exactly as an unexpanded field would be.

## MergeRange

| Field     | Type     | Description                     |
| --------- | -------- | ------------------------------- |
| `row`     | `number` | Start row (0 = first data row)  |
| `col`     | `number` | Start column (0 = first column) |
| `rowspan` | `number` | Row span                        |
| `colspan` | `number` | Column span                     |

## CellStyle

| Field        | Type                                                   | Description                                  |
| ------------ | ------------------------------------------------------ | -------------------------------------------- |
| `font?`      | `{ bold?, italic?, size?, color?, name? }`             | `color` is 6-digit RGB hex (e.g. `"FF0000"`) |
| `fill?`      | `{ pattern?: "solid" \| "none", fgColor?, bgColor? }`  | Fill                                         |
| `alignment?` | `{ horizontal?, vertical?, wrapText?, textRotation? }` | Alignment (textRotation 0–180)               |
| `border?`    | `{ top?, bottom?, left?, right? }`                     | Borders, each `{ style, color? }`            |
| `numFormat?` | `string`                                               | Excel number format code                     |

## ExportMode / ExportPhase

```ts
type ExportMode = "auto" | "main" | "worker" | "stream";
type ExportPhase = "init" | "build" | "download";
```

## Full import

```ts
import type {
  SheetConfig,
  ColumnConfig,
  CellStyle,
  MergeRange,
  FormatSpec,
  ExportOptions,
  ExportResult,
  ExportMode,
  ExportPhase,
  BorderStyle,
} from "@marcusok/excel-exporter";
```
