# API: Core Types

## SheetConfig

| Field          | Type                            | Required | Description                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| -------------- | ------------------------------- | -------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `name`         | `string`                        | yes      | Non-empty, ≤ 31 chars, no `: \ / ? * [ ]`, no leading/trailing apostrophe. Uniqueness across `sheets` is **case-insensitive** (Excel's own rule — `Sheet1` and `SHEET1` clash and are rejected)                                                                                                                                                                                                                                                                            |
| `columns`      | `ColumnConfig[]`                | yes      | Column definitions                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| `data`         | `Record<string, unknown>[]`     | yes      | Row data. Without a column `format`, values are normalized identically on every route (main / worker / stream, fallback included): `Date` → ISO string, plain objects → JSON, `NaN`/`Infinity`/`bigint` → visible strings — so a dataset crossing the 50k threshold keeps the same cell content. A raw `Date` without a `date`/`datetime` format therefore lands as **text**, not an Excel date (see the [formatting guide](/packages/excel-exporter/guide/04-formatting)) |
| `headerStyle?` | `CellStyle`                     | —        | Default header style; overridden by column-level `headerStyle`. Dropped on the stream path (see the note below)                                                                                                                                                                                                                                                                                                                                                            |
| `dataStyle?`   | `CellStyle`                     | —        | Base style for every data cell; column `style` deep-merges over it field by field (see the Styles guide). Dropped on the stream path                                                                                                                                                                                                                                                                                                                                       |
| `indexColumn?` | `boolean \| IndexColumnOptions` | —        | Inject a leading row-number column; `true` = all defaults. Merges shift right automatically. Works on every route, stream included                                                                                                                                                                                                                                                                                                                                         |
| `freezeRows?`  | `number`                        | —        | Number of worksheet rows frozen from the top — usually the header row count, but not tied to it (freezing 5 with a 2-row header also freezes 3 data rows); validated as a non-negative integer. Dropped on the stream path                                                                                                                                                                                                                                                 |
| `merges?`      | `MergeRange[]`                  | —        | Merged cells (relative to the data area). Preserved on the stream path (merges are structure, not styles)                                                                                                                                                                                                                                                                                                                                                                  |
| `autoFilter?`  | `boolean`                       | —        | Auto filter over the last header row + all data rows. Dropped on the stream path                                                                                                                                                                                                                                                                                                                                                                                           |

## ColumnConfig

| Field          | Type                     | Required     | Description                                                                                                                                          |
| -------------- | ------------------------ | ------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| `prop?`        | `string`                 | leaf columns | Field name on the data row (Element Plus naming); group columns (with `children`) may omit it                                                        |
| `key?`         | `string`                 | —            | Deprecated alias of `prop` (pre-2.2 naming); `prop` wins when both are present                                                                       |
| `label?`       | `string`                 | yes*         | Header text (leaf and group columns alike); `label` or legacy `header` — at least one required (*validated at export time)                           |
| `header?`      | `string`                 | —            | Deprecated alias of `label` (pre-2.2 naming); `label` wins when both are present                                                                     |
| `children?`    | `ColumnConfig[]`         | —            | Group column: produces a multi-row header; its header cell merges across all descendant leaf columns. `children: []` is a leaf                       |
| `width?`       | `number`                 | —            | Column width (Excel character units; `0` hides the column); validated as a finite non-negative number; leaf columns only. Dropped on the stream path |
| `style?`       | `CellStyle`              | —            | Data-cell style (headers excluded); leaf columns only. Dropped on the stream path                                                                    |
| `headerStyle?` | `CellStyle`              | —            | Header style for this column (group header cells included); wins over sheet-level `headerStyle`. Dropped on the stream path                          |
| `format?`      | `FormatSpec \| Function` | —            | Value formatting; leaf columns only; functions run on main-thread paths and are stripped on the browser worker path (see the FormatSpec page)        |

A column with `children` is a group: no data cells, header rows only. Header row count = 1 + the deepest column's tree depth (so 1 for a flat column list); leaf headers span the remaining header rows vertically, group headers span their leaf subtree horizontally — merges are generated automatically (no manual `merges` needed for headers).

Three structural validations run up front: an empty `columns` array is rejected (`sheet has no columns`); the same column object appearing more than once in the tree — e.g. one object shared as a child of two groups — is rejected (`column object reused in the column tree`); and a `children` reference cycle is rejected (`circular children reference in column tree`).

Sheet dimensions are validated against Excel's hard limits up front as well: 1,048,576 rows per sheet (header rows count toward it) and 16,384 columns (an injected `indexColumn` occupies one). An oversized sheet resolves `{ success: false }` with an error naming the sheet and its actual row/column counts — on the stream path it would otherwise zip successfully and only surface as Excel's repair prompt on open.

**Stream-path compatibility.** On the fast stream — explicit `mode: "stream"`, auto at/above the row or cell gate, or the degraded fallback — cell/layout styling is dropped with a `console.warn` and reported as `ExportResult.degradedFeatures`: `width`, column `style` / `headerStyle` (including `style.numFormat`), sheet-level `headerStyle` / `dataStyle`, `freezeRows`, `autoFilter`. `FormatSpec` date/number numFmts are **not** dropped: the stream emits a minimal `styles.xml`, so dates stay serial numbers and decimals/thousands are display-only on both paths. Structure survives: multi-row headers, `merges` and `indexColumn` are preserved. The full per-route matrix lives in the [worker & stream guide](/packages/excel-exporter/guide/06-worker-stream).

## IndexColumnOptions

Options for `SheetConfig.indexColumn`; the shorthand `true` equals `{}`.

| Field          | Type        | Default  | Description                                                                                      |
| -------------- | ----------- | -------- | ------------------------------------------------------------------------------------------------ |
| `label?`       | `string`    | `"序号"` | Header text of the index column                                                                  |
| `width?`       | `number`    | `6`      | Column width (Excel character units; `0` hides the column)                                       |
| `start?`       | `number`    | `1`      | Number shown on the first data row; row i displays `start + i`. Non-negative integer (validated) |
| `style?`       | `CellStyle` | —        | Data-cell style of the index column; merged over sheet-level `dataStyle` like any column style   |
| `headerStyle?` | `CellStyle` | —        | Header style of the index column; overrides sheet-level `headerStyle`                            |

The index column's values are generated from the row number and never read from `data`. Through `exportExcel` / `applyIndexColumn`, a user column declaring the reserved `__index__` prop is rejected with a clear error; the lower-level builders do not perform that check — a hand-written `__index__` (or legacy `key`) leaf is treated as the row-number column there, so its data values are replaced by row numbers. Existing `merges` are shifted one column right so they keep pointing at their original targets.

`exportExcel` expands this field automatically (once, before mode routing), which is what makes it work on every route. The lower-level entry points — `WorkbookBuilder.addSheet()` and `exportAsStream()` — do **not** expand it: they only understand an already-expanded `__index__` column, so pass `applyIndexColumn(sheet)` (exported from the package, along with the reserved `INDEX_PROP`) when calling them directly. Otherwise `indexColumn` is silently ignored, exactly as an unexpanded field would be. The `WorkbookBuilder` direct path also does not enforce the case-insensitive duplicate-name rule from the `name` row above: it only gets the engine's exact-case duplicate check, so `Report` / `REPORT` can both be added. (`exportAsStream` performs the same lowercase duplicate check as `exportExcel`.)

## MergeRange

| Field     | Type     | Description                     |
| --------- | -------- | ------------------------------- |
| `row`     | `number` | Start row (0 = first data row)  |
| `col`     | `number` | Start column (0 = first column) |
| `rowspan` | `number` | Row span                        |
| `colspan` | `number` | Column span                     |

Validation runs up front, identically on every path, with a message naming the offending merge: all four values must be integers; `row`/`col` ≥ 0 and `rowspan`/`colspan` ≥ 1; `col + colspan` must not exceed the sheet's leaf-column count and `row + rowspan` must not exceed its data-row count (header rows live outside this coordinate space); ranges must not overlap each other (they sit entirely below the header block, so they cannot collide with the generated header merges). A non-array `merges` is rejected as well.

## CellStyle

| Field        | Type                                                   | Description                                                                                                                                              |
| ------------ | ------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `font?`      | `{ bold?, italic?, size?, color?, name? }`             | `color` is 6-digit RGB hex (e.g. `"FF0000"`)                                                                                                             |
| `fill?`      | `{ pattern?: "solid" \| "none", fgColor?, bgColor? }`  | Fill                                                                                                                                                     |
| `alignment?` | `{ horizontal?, vertical?, wrapText?, textRotation? }` | Alignment: `horizontal` `"left" \| "center" \| "right"`, `vertical` `"top" \| "center" \| "bottom"`, `wrapText` boolean, `textRotation` an integer 0–180 |
| `border?`    | `{ top?, bottom?, left?, right? }`                     | Borders, each `{ style: BorderStyle, color? }` — see [BorderStyle](#borderstyle)                                                                         |
| `numFormat?` | `string`                                               | Excel number format code                                                                                                                                 |

Style fields are validated up front — an invalid value fails the export with `{ success: false }` and a field-specific message (naming the sheet/column) instead of surfacing later as an opaque engine error or a degraded export: colors must be 6-digit RGB hex (`"FF0000"`; `"red"` and `"#FF0000"` are both rejected), `border.*.style` must be one of the `BorderStyle` names above, `font.size` a finite positive number, and `alignment.textRotation` an **integer** in 0–180 (`45.5` is rejected, not rounded). The checks cover sheet-level, column-level (group columns' ignored styles included) and `indexColumn` styles alike. They also cover `fill.pattern` (`"solid"` / `"none"`), the alignment horizontal/vertical enums, `wrapText`/`bold`/`italic` types, and the types of `font.name` / `numFormat`. The latter two, plus sheet names and date/datetime patterns, reject XML 1.0-illegal characters (control characters other than tab/LF/CR, U+FFFE/U+FFFF, lone surrogates) that the engine would otherwise write verbatim and Excel would treat as a corrupt file. Data/header strings are sanitized instead (illegal controls removed, lone surrogates replaced with U+FFFD) identically on every route.

## BorderStyle

The `style` value of each border side — Excel's own line-style vocabulary, inlined from modern-xlsx's `BorderSideData` so the published `.d.ts` carries no dependency imports:

```ts
type BorderStyle =
  | "thin"
  | "medium"
  | "thick"
  | "dashed"
  | "dotted"
  | "double"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot";
```

Each name carries its own weight — there is no separate width knob (`hair` is the finest line, `thick` the heaviest). The preview's renderer maps these same values onto CSS; the Excel → CSS table lives in the [preview's data model reference](/packages/excel-preview/api/02-model#borders-previewborder).

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
