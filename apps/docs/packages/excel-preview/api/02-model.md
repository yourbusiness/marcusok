# Data Model (PreviewWorkbook)

`parseWorkbookBytes` (and the worker) emit a plain JSON model — structured-clone safe, framework-neutral, and exactly what the built-in renderer consumes. This page is the field-level reference: for every field it documents the value the parser actually produces and what the built-in renderer does with it, so a renderer can be written against the model alone.

## Shape

```ts
interface PreviewWorkbook {
  sheets: PreviewSheet[];
  activeSheetIndex: number; // 0-based, from the file's activeTab (clamped to 0)
  dateSystem: "date1900" | "date1904"; // drives serial→date conversion
}

interface PreviewSheet {
  name: string;
  visible: boolean; // hidden / veryHidden sheets stay in the list
  showGridLines: boolean; // from the file's view settings; default true
  rightToLeft: boolean;
  rowCount: number; // 1-based data bounds — see Bounds below
  colCount: number;
  rows: PreviewRow[]; // sparse: see Cells below
  colSpans: PreviewColSpan[]; // from the file's <col> elements, 1-based min/max
  merges: PreviewMerge[]; // 0-based { row, col, rowSpan, colSpan }
  frozenRows: number; // 0 = none
  frozenCols: number;
  styles: PreviewStyles; // workbook-level: shared by every sheet
}

interface PreviewRow {
  index: number; // 1-based
  height: number | null; // pt; null → default (15pt)
  hidden: boolean;
  cells: PreviewCell[]; // sparse by column, ordered by col
}

interface PreviewCell {
  col: number; // 0-based
  type: PreviewCellType; // "number" | "string" | "boolean" | "error" | "formulaStr"
  value: string | null; // raw string form; numbers keep full precision
  styleIndex: number | null; // index into styles.xfs; null = default style
}

interface PreviewColSpan {
  min: number; // 1-based, inclusive
  max: number; // 1-based, inclusive
  width: number; // Excel character units; 8.43 when the file omits it
  hidden: boolean; // hidden columns are present in the model (width renders as 0)
  customWidth: boolean; // metadata only — see Column spans below
}

interface PreviewMerge {
  row: number; // 0-based anchor (top-left)
  col: number;
  rowSpan: number; // >= 1
  colSpan: number;
}

interface PreviewStyles {
  fonts: PreviewFont[];
  fills: PreviewFill[];
  borders: PreviewBorder[];
  xfs: PreviewXf[]; // cell styles, indexed by PreviewCell.styleIndex
}
```

## Bounds (`rowCount` / `colCount`)

Both are 1-based **data bounds**, not array lengths — and they are deliberately allowed to be larger than anything the `rows` array suggests:

| Bound      | Computed as the maximum of                                                                       |
| ---------- | ------------------------------------------------------------------------------------------------ |
| `rowCount` | the largest `PreviewRow.index` present, and `row + rowSpan` of every merge                       |
| `colCount` | `col + 1` for every cell, `PreviewColSpan.max` of every span, and `col + colSpan` of every merge |

Two consequences that a renderer must not miss:

- **Merges push the bounds outward.** A sheet whose only value sits in `A1` with a merge `A1:C3` reports `rowCount: 3, colCount: 3`, even though `rows` has a single entry and no cell exists at column 2 or 3. Excel does not write covering cells for a merge, so the bounds are the only record of the merged block's extent — size your layout from them or the block gets truncated to 1×1.
- **`<col>` spans push `colCount` outward.** A file declaring `<col min="1" max="20">` reports `colCount >= 20` whether or not those columns hold any content.

`0` is possible for a genuinely empty sheet (no rows, no spans, no merges). The built-in layout additionally floors the extent at the frozen row/column counts, and at `1`: a freeze pane may declare more frozen rows than the sheet has content.

## Cells

**`type` is a normalized value, not the engine's `cellType`.** The parser collapses the engine's string-producing variants into one:

| Engine `cellType`                                     | Model `type`   |
| ----------------------------------------------------- | -------------- |
| `number`                                              | `"number"`     |
| `boolean`                                             | `"boolean"`    |
| `error`                                               | `"error"`      |
| `formulaStr`                                          | `"formulaStr"` |
| `sharedString` / `inlineStr` / `stub` / anything else | `"string"`     |

So the model never records where a string came from — a shared-string-table entry and an inline string are both `"string"` — and an unrecognized engine type degrades to `"string"` rather than failing. Note that `"stub"` (a referenced-but-empty shared string) merges into `"string"` as well; those cells normally disappear earlier, because a cell is only emitted when it has a value **or** a style.

**`value` is always the raw string form.** Numbers keep full precision (`"45678.5"`), booleans arrive in the engine's raw form (`"1"` / `"0"` on most files), `error` carries the error literal (`"#DIV/0!"`), and `formulaStr` carries the formula's cached string result. A formula with a numeric cached result arrives as `type: "number"` with that cached value — matching how Excel displays it.

**`value: null` means "formatted but empty"**, not "absent": a cell with no value but a style is still emitted so its fill and borders render (format it to an empty string). A cell with neither a value nor a style is dropped entirely.

**`styleIndex` is an index into `styles.xfs`; `null` means the default style.** Do not index `xfs[0]` for it — see **Formatting a cell yourself** below.

`type` also drives the built-in renderer's General alignment rule: numbers right, strings and `formulaStr` left, booleans and errors centered. That only applies when the xf carries no explicit horizontal alignment.

## Column spans

`colSpans` mirrors the file's `<col>` elements one-to-one, in document order — it is **not** expanded per column, so a single entry may cover hundreds of columns.

- `min` / `max` are 1-based and inclusive: `{ min: 2, max: 5 }` covers columns B–E (0-based 1–4).
- `width` is in Excel character units, the unit of the file's own `width` attribute. When the file omits it the parser falls back to `8.43` — which is Excel's default column width, so the fallback is visually a no-op.
- `hidden` columns keep their slot with a width of **0** (they are not removed from the coordinate space; see **Writing a custom renderer** below).
- `customWidth` is retained as **metadata only**. Do not gate the width on it: the engine reports `customWidth: false` for both forms in practice, so gating silently drops the column widths written by third-party producers that emit `width` without the flag. Any `<col>` present contributes its width.

To convert a width to pixels the way the built-in renderer does: `round(width * 7 + 5)` (Calibri 11 at 96dpi, max digit width 7). Column-level styles (`<col style="…">`) are not carried in the model.

## Styles

`PreviewStyles` is workbook-level and shared: the file has one `styles.xml`, so every sheet points at the same object. Do not deep-copy it per sheet — build your stylesheet once.

### Cell styles (`PreviewXf`)

| Field        | Type                    | Meaning                                                            |
| ------------ | ----------------------- | ------------------------------------------------------------------ |
| `fontId`     | `number`                | Index into `styles.fonts`                                          |
| `fillId`     | `number`                | Index into `styles.fills`                                          |
| `borderId`   | `number`                | Index into `styles.borders`                                        |
| `numFmtCode` | `string`                | Resolved format code; never `undefined` (`"General"` is the floor) |
| `alignment`  | `AlignmentData \| null` | Explicit alignment, or `null` for none                             |

The id fields come straight from the file's `cellXfs` (floored to `0` when the attribute is missing), while the three collections are rebuilt in `styles.xml` document order so the ids line up — self-closed default entries included as empty items. Index defensively anyway: an out-of-range id yields `undefined`, and the built-in compiler then emits nothing for that slot.

### Fonts (`PreviewFont`)

| Field                                      | Type             | Meaning                                                                         |
| ------------------------------------------ | ---------------- | ------------------------------------------------------------------------------- |
| `name`                                     | `string \| null` | Font family name; `null` → the renderer's own stack (Calibri, Segoe UI…)        |
| `size`                                     | `number \| null` | Size in **pt**; **`null` = the default 11pt**                                   |
| `bold` / `italic` / `underline` / `strike` | `boolean`        | One flag per attribute; underline and strike combine into one `text-decoration` |
| `color`                                    | `string \| null` | Resolved CSS color; **`null` = automatic → rendered black**                     |

A hostile file's font name is stripped of `"` and `\` before it is interpolated into a CSS `font-family` declaration (an unescaped quote would otherwise terminate the generated rule and let the file inject arbitrary declarations). Every other interpolated value is a validated hex color or a fixed keyword.

### Fills (`PreviewFill`)

A discriminated union on `kind` — the four variants have different shapes, so narrow first:

```ts
type PreviewFill =
  | { kind: "none" }
  | { kind: "solid"; fgColor: string | null }
  | { kind: "pattern"; fgColor: string | null; bgColor: string | null }
  | {
      kind: "gradient";
      degree: number;
      stops: { position: number; color: string }[];
    };
```

| `kind`     | Produced when                                    | How the built-in renderer draws it                                                                               |
| ---------- | ------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------- |
| `none`     | No fill, or the pattern type is missing / `none` | Nothing                                                                                                          |
| `solid`    | Pattern type `solid`                             | `background-color: fgColor`; a `null` `fgColor` (automatic) is treated as **no fill** rather than painting black |
| `pattern`  | Any other pattern type                           | Approximated as a solid: `background-color: bgColor ?? fgColor ?? #ffffff` — the hatch itself is not reproduced  |
| `gradient` | A `gradientFill` with at least one stop          | `linear-gradient((degree + 90) % 360deg, …)`, stop positions as percentages                                      |

Gradient details worth knowing when re-implementing:

- `position` is `0..1`. The parser fills in missing positions with an even spread (`i / (n - 1)`) and sorts the stops by position; a stop with no resolvable color falls back to `#000000`.
- `degree` follows ECMA-376: `0` = left→right, `90` = top→bottom. CSS `0deg` points up and `90deg` points right, hence the `+ 90` shift.
- Gradient fills and diagonal borders both have to be emitted as `background-image`. Two separate declarations do not stack — the later one wins — so they must be combined into one comma-separated multi-layer declaration (diagonal on top).

### Borders (`PreviewBorder`)

Five sides, each either `{ style: BorderStyle; color: string | null }` or `null`:

| Side                                | Meaning                                                           |
| ----------------------------------- | ----------------------------------------------------------------- |
| `left` / `right` / `top` / `bottom` | An edge of the cell box; `null` = no border on that edge          |
| `diagonal`                          | The cell's diagonal line; `null` = none (see the flag rule below) |

A side is `null` when the file declares no style for it or declares `style="none"`. `color: null` is an automatic color and is painted `#000000`.

The four edge sides map onto CSS `border-<side>` declarations:

| Excel `style`                                                       | CSS emitted by the built-in renderer |
| ------------------------------------------------------------------- | ------------------------------------ |
| `thin`, `hair`                                                      | `1px solid`                          |
| `medium`                                                            | `2px solid`                          |
| `thick`                                                             | `3px solid`                          |
| `double`                                                            | `3px double`                         |
| `dashed`, `dashDot`, `dashDotDot`                                   | `1px dashed`                         |
| `mediumDashDot`, `mediumDashDotDot`, `mediumDashed`, `slantDashDot` | `2px dashed`                         |
| anything unrecognized                                               | `1px solid`                          |

CSS has no native diagonal, so `diagonal` is drawn as a thin linear-gradient layer across the cell box instead of a border (2px wide for `medium` / `thick`, 1px otherwise). Its direction is a separate, optional field:

- **`diagonalUp` is the direction flag: `true` = `/` (Excel's `diagonalUp`); absent or `false` = `\` (Excel's `diagonalDown`, the more common form).**
- **A file that declares a diagonal style without either the up or the down flag renders with no diagonal at all**, because Excel itself does not draw one. The parser enforces this by setting `diagonal: null` in that case, so the flag never has to be second-guessed at render time — and `diagonalUp` is only meaningful when `diagonal` is non-null.
- When both flags are set (an X shape) a single gradient cannot express two lines, so the down form wins and the result is `\`.
- One exception to the flag rule: when `styles.xml` cannot be read at all, the parser falls back to the engine's border arrays, which do not carry the flags. That degraded path keeps whatever diagonal the engine returned and pins `diagonalUp: false` (`\`) — it is a guess either way, and it only applies to files whose styles part is unreadable.

### Alignment (`PreviewXf.alignment`)

The type is `AlignmentData | null`, re-exported unchanged from `@marcusok/xlsx-core` (the parser does not re-declare it). `null` means the file set no alignment, so the renderer's by-type General rule applies. The engine type declares six fields, but **the renderer consumes five of them**:

| Field          | Type                                                                                                               | Consumed | Renderer behavior                                                                                                                                                                                                                                                                    |
| -------------- | ------------------------------------------------------------------------------------------------------------------ | -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `horizontal`   | `"general" \| "left" \| "center" \| "right" \| "fill" \| "justify" \| "centerContinuous" \| "distributed" \| null` | yes      | `left` / `center` / `right` map onto the flex main axis; `justify` and `distributed` keep the text left-aligned and add `text-align: justify`; `general`, `fill` and `centerContinuous` fall through to the by-type rule (`fill` and `centerContinuous` are approximated as General) |
| `vertical`     | `"top" \| "center" \| "bottom" \| "justify" \| "distributed" \| null`                                              | yes      | `top` / `center` / `bottom` map onto the flex cross axis; `justify` and `distributed` have no CSS equivalent and are approximated as `center`                                                                                                                                        |
| `wrapText`     | `boolean`                                                                                                          | yes      | Adds `white-space: pre-wrap` + `word-break: break-word`                                                                                                                                                                                                                              |
| `indent`       | `number \| null`                                                                                                   | yes      | See below                                                                                                                                                                                                                                                                            |
| `textRotation` | `number \| null`                                                                                                   | yes      | See below                                                                                                                                                                                                                                                                            |
| `shrinkToFit`  | `boolean`                                                                                                          | **no**   | Declared by the engine type but ignored: text is not scaled down to fit; overflow follows the normal clip/spill rules                                                                                                                                                                |

`indent` conversion: Excel's indent unit is roughly three space characters wide, so each level becomes `padding-inline-start: indent * 3ch` (the `ch` unit is relative to the cell's own font size, which is what makes it a good proxy). For `horizontal: "right"` the padding is moved to the end of the line instead, since Excel indents right-aligned text from the right.

`textRotation` mapping:

| Value       | Renderer behavior                                                                                      |
| ----------- | ------------------------------------------------------------------------------------------------------ |
| `null`, `0` | No rotation                                                                                            |
| `1`–`90`    | Counter-clockwise: `rotate(-n deg)`, origin at the bottom-left, text kept on one line                  |
| `91`–`180`  | Clockwise: `rotate((n - 90) deg)`, origin at the bottom-right, text kept on one line                   |
| `255`       | Vertically stacked text (`writing-mode: vertical-rl`) — Excel's "vertical text" option, not a rotation |

Rotation cannot be reflected in the cell box, so rotated runs stay on a single line and are clipped by the cell rectangle when they overflow — see [Scope & Limits](/packages/excel-preview/guide/04-limits) for how that differs from Excel near the box diagonal.

## Key points for consumers

- **Values are raw strings.** Numbers keep their full precision as strings; use `Number(value)` when you need the number. The renderer does exactly that on its formatting path.
- **`numFmtCode` is pre-resolved.** Builtin ids were mapped through the Excel-behavior table (id 14 → `m/d/yyyy`), custom ids came from the file's `numFmts`. It is never `undefined` — `"General"` is the floor.
- **Colors are CSS strings** (`#rrggbb` / `rgba(...)`), already resolved through the theme/indexed overlay; `null` means automatic (render defaults: black text, no fill).
- **`formulaStr` cells** carry a formula's cached string result; formula cells with a numeric cached value come through as `type: "number"` with the cached value (matching how Excel displays them).
- **Empty cells are absent** — the model is sparse on both axes; renderers should treat absence as an empty cell with default style.
- **The model is pure JSON**, with no class instances or `Date` objects, so it survives `structuredClone` and can be cached or transferred between contexts as-is. It is also read-only by contract: derive your own indexes rather than annotating it.

## Formatting a cell yourself

The exact formatter the renderer uses is exported for wrappers that build their own DOM:

```ts
import { formatCellValue } from "@marcusok/excel-preview";

// Mirrors the renderer: no xf means no format code, so "General" — do not
// substitute xfs[0], which merely happens to be the Normal style in most files.
const xf =
  cell.styleIndex != null ? sheet.styles.xfs[cell.styleIndex] : undefined;

const { text, color } = formatCellValue(
  cell.type,
  cell.value,
  xf?.numFmtCode ?? "General",
  workbook.dateSystem,
);
// text: the display string; color: "#ff0000" when the format has a [Red]-style section
```

| Parameter    | Type                       | Notes                                                                                                                     |
| ------------ | -------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `type`       | `PreviewCellType`          | Pass `cell.type` through: the formatter branches on it (booleans → `TRUE`/`FALSE`, errors verbatim, General numbers)      |
| `value`      | `string \| null`           | Pass `cell.value`; `null` or `""` returns `{ text: "" }`                                                                  |
| `numFmtCode` | `string`                   | The resolved code for the cell's xf; an empty or missing code is treated as `"General"`, and the call never throws for it |
| `dateSystem` | `"date1900" \| "date1904"` | `workbook.dateSystem`; 1904 serials are shifted +1462 days before date formatting                                         |

It returns `FormattedValue` (exported as a type from the package):

```ts
interface FormattedValue {
  text: string; // the display string
  color?: string; // CSS color from a [Red] / [Color 3] section, when present
}
```

`color` is optional and only set when the chosen format section carries a color bracket. Apply it as an inline color that overrides the xf's font color (the renderer sets `el.style.color`), and only when present — overwriting with `undefined` would drop a legitimate black.

What the formatter restores and what it approximates is documented in full on [Format Fidelity](/packages/excel-preview/guide/03-format-fidelity).

## Writing a custom renderer

Assuming you build your own DOM from this model, these are the properties that bite:

1. **The model is sparse in both directions.** `rows` holds only rows that have content, a custom height or a hidden flag; `cells` holds only columns that have a value or a style. Iterate using the `index` / `col` fields, never an array position — rows are sorted by `index`, but the gaps are real.
2. **Absent is not the same as "empty but styled".** A missing `(row, col)` pair is a blank cell with the default style: draw your grid background and grid lines, not a default-styled content box. A cell present with `value: null` is a formatted empty cell and must still render its fill and borders.
3. **`rowCount` / `colCount` size the grid, not the arrays.** They include the extent of merges and `<col>` spans (see **Bounds** above), so they are the only correct source for the scrollable extent.
4. **Numbers are strings.** Convert with `Number(value)`; never re-parse a display string. Keeping the raw string is what preserves precision that a `Number` round-trip could lose on the way out of the file.
5. **Colors are already resolved CSS strings.** Use them verbatim — `#rrggbb` for explicit RGB, `rgba(...)` for themed colors carrying a tint. `null` means automatic: black for fonts and borders, nothing painted for fills.
6. **`styleIndex: null` means the default style**, not `xfs[0]`. Skip the class and the xf lookup entirely, and format with `"General"`.
7. **Compile styles once per workbook, not per cell.** The built-in renderer emits one CSS class per xf index and gives each cell just its class name; per-cell inline style objects would be recomputed on every virtual-scroll rebuild, while a class lookup is free and cacheable by the browser. Only the format-code color (from `FormattedValue.color`) needs an inline style.
8. **Skip the cells a merge covers.** A merge is one entry anchored at its top-left; the covered pairs are absent from `rows`. Render the anchor box at the merged size (the sum of the spanned column widths and row heights) and let it cover the interior — including anchors whose top-left is just outside the viewport, so that a merged block crossing the viewport is drawn whole.
9. **Hidden rows and columns keep their slot with size 0.** They still occupy their index in the coordinate space; dropping them from your prefix sums shifts every following row or column. A `<col hidden>` is a column of width 0, not a missing column.
10. **Treat the model as read-only.** It arrives by structured clone from the worker and may be reused across renders; build your own lookup structures (for example a `Map` keyed by `row:col`) instead of mutating it.
