# Data Model (PreviewWorkbook)

`parseWorkbookBytes` (and the worker) emit a plain JSON model — structured-clone safe, framework-neutral, and exactly what the built-in renderer consumes.

## Shape

```ts
interface PreviewWorkbook {
  sheets: PreviewSheet[];
  activeSheetIndex: number; // 0-based, from the file's activeTab
  dateSystem: "date1900" | "date1904"; // drives serial→date conversion
}

interface PreviewSheet {
  name: string;
  visible: boolean; // hidden / veryHidden sheets stay in the list
  showGridLines: boolean;
  rightToLeft: boolean;
  rowCount: number; // 1-based data bounds
  colCount: number;
  rows: PreviewRow[]; // sparse: rows without content/height/hidden are absent
  colSpans: PreviewColSpan[]; // <col> spans (width/hidden/customWidth), 1-based min/max
  merges: PreviewMerge[]; // 0-based { row, col, rowSpan, colSpan }
  frozenRows: number; // 0 = none
  frozenCols: number;
  styles: PreviewStyles; // shared across sheets (workbook-level styles.xml)
}

interface PreviewRow {
  index: number; // 1-based
  height: number | null; // pt; null → default (15pt)
  hidden: boolean;
  cells: PreviewCell[];
}

interface PreviewCell {
  col: number; // 0-based
  type: "number" | "string" | "boolean" | "error" | "formulaStr";
  value: string | null; // numbers stay strings ("45678.5"); null = empty
  styleIndex: number | null; // index into styles.xfs; null = default
}

interface PreviewStyles {
  fonts: PreviewFont[]; // { name, size(pt), bold, italic, underline, strike, color }
  fills: PreviewFill[]; // none | solid | pattern | gradient (stops resolved)
  borders: PreviewBorder[]; // per-side { style, color } | null
  xfs: PreviewXf[]; // { fontId, fillId, borderId, numFmtCode, alignment }
}
```

Key points for consumers:

- **Values are raw strings.** Numbers keep their full precision as strings; use `Number(value)` when you need the number. The renderer does exactly that on its formatting path.
- **`numFmtCode` is pre-resolved.** Builtin ids were mapped through the Excel-behavior table (id 14 → `m/d/yyyy`), custom ids came from the file's `numFmts`. It is never `undefined` — `"General"` is the floor.
- **Colors are CSS strings** (`#rrggbb` / `rgba(...)`), already resolved through the theme/indexed overlay; `null` means automatic (render defaults: black text, no fill).
- **`formulaStr` cells** carry a formula's cached string result; formula cells with a numeric cached value come through as `type: "number"` with the cached value (matching how Excel displays them).
- **Empty cells are absent** — the model is sparse on both axes; renderers should treat absence as an empty cell with default style.

## Formatting a cell yourself

The exact formatter the renderer uses is exported for wrappers that build their own DOM:

```ts
import { formatCellValue } from "@marcusok/excel-preview";

const { text, color } = formatCellValue(
  cell.type,
  cell.value,
  sheet.styles.xfs[cell.styleIndex ?? 0]?.numFmtCode ?? "General",
  workbook.dateSystem,
);
// text: display string; color: "#ff0000" when the format has a [Red]-style section
```
