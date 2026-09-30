# Engine Surface

Everything below is a **re-export** of [modern-xlsx](https://github.com/ABCrimson/modern-xlsx)'s public API — the subset the @marcusok packages share, plus the types they need. The signatures here are the engine's own; this package adds no wrapper, so anything the engine documents applies unchanged.

## Workbook, reading and writing

```ts
import {
  Workbook,
  readBuffer,
  initWasm,
  initWasmSync,
} from "@marcusok/xlsx-core";

function readBuffer(data: Uint8Array, options?: ReadOptions): Promise<Workbook>;
class Workbook {
  constructor(data?: Partial<WorkbookData>);
  readonly sheetNames: readonly string[];
  getSheet(name: string): Worksheet | undefined;
  addSheet(name: string): Worksheet;
  toBuffer(options?: WriteOptions): Promise<Uint8Array>;
}
function initWasm(wasmSource?: string | URL | Response): Promise<void>;
function initWasmSync(module: WebAssembly.Module | BufferSource): void;
```

```ts
const workbook = await readBuffer(bytes); // ReadOptions = { password?: string }
const sheet = workbook.getSheet("Sheet1") ?? workbook.addSheet("Sheet1");
sheet.cell("A1").value = "Hello";
const out = await workbook.toBuffer(); // WriteOptions = { password?: string }
```

`initWasm` / `initWasmSync` are the engine's own initialization calls. You rarely need them: `getWasmLoader().ensureLoaded()` runs them for you ([see the loader guide](/packages/xlsx-core/guide/01-loader)). The exception is a **self-contained worker entrypoint**, which cannot use the loader's Node auto-init path and therefore calls `initWasm(wasmUrl)` with an explicit URL — exactly what `@marcusok/excel-preview`'s parse worker does. `initWasmSync` is what the loader's Node auto-init calls internally.

## Cell and range references

```ts
import {
  encodeCellRef,
  decodeCellRef,
  encodeRange,
  decodeRange,
  columnToLetter,
  letterToColumn,
} from "@marcusok/xlsx-core";

function encodeCellRef(row: number, col: number): string; // (0, 0) → "A1"
function decodeCellRef(ref: string): CellAddress; // "A1" → { row: 0, col: 0 }
function encodeRange(start: CellAddress, end: CellAddress): string; // → "A1:C10"
function decodeRange(range: string): CellRange; // → { start: CellAddress, end: CellAddress }
function columnToLetter(col: number): string; // 0 → "A", 26 → "AA"
function letterToColumn(letter: string): number; // "A" → 0, "AA" → 26
```

All indices are **0-based**; the A1 strings are 1-based, which is why `encodeCellRef(0, 0)` is `"A1"`. `CellAddress` is `{ readonly row: number; readonly col: number }`, `CellRange` is `{ readonly start: CellAddress; readonly end: CellAddress }`.

This is the group the exporter leans on hardest: it builds A1 references for the cells it writes (`encodeCellRef`) and parses the merge ranges it is asked to write (`decodeCellRef`). `columnToLetter` / `letterToColumn` are not used by either business package today — they are here for renderers that draw their own column headers.

## Number formatting

```ts
import {
  formatCell,
  formatCellRich,
  getBuiltinFormat,
  loadFormatTable,
  isDateFormatId,
  isDateFormatCode,
} from "@marcusok/xlsx-core";

function formatCell(
  value: string | number | boolean | null,
  format: string | number, // format code, or a built-in format id
  opts?: FormatCellOptions, // { dateSystem?: DateSystem }
): string;

function formatCellRich(
  value: string | number | boolean | null,
  format: string | number,
  opts?: FormatCellOptions,
): FormatCellResult; // { text: string; color?: string }

function getBuiltinFormat(id: number): string | undefined;
function loadFormatTable(table: Record<number, string>): void;
function isDateFormatId(numFmtId: number): boolean;
function isDateFormatCode(formatCode: string): boolean;
```

`formatCellRich` additionally supports conditional sections (`[>100]#,##0;[<=100]0.00`) and bracket color directives (`[Red]`, `[Color3]`), returning the color name alongside the text — this is the function behind the preview's formatting layer:

```ts
const { text, color } = formatCellRich(1234.5, "#,##0.00");
// text: "1,234.50"
```

Caveats that are the engine's, not this package's: `formatCell` / `formatCellRich` throw a `TypeError` on a `null` / `undefined` format code (always resolve an id to a code string first), and the engine drops literals inside numeric sections of multi-section formats. The preview ships a compensation layer over exactly these defects — read [Format Fidelity](/packages/excel-preview/guide/03-format-fidelity) before you rely on `formatCellRich` output for display parity with Excel.

`isDateFormatId` only consults the built-in ids; for custom codes use `isDateFormatCode`, which scans the code string for date/time tokens while ignoring quoted strings. Neither is used by the two business packages today — they are part of the surface because any renderer that decides "is this number a date?" needs them.

## Date serials

```ts
import { serialToDate, dateToSerial } from "@marcusok/xlsx-core";

function dateToSerial(date: Date | TemporalLike, system?: DateSystem): number;
function serialToDate(serial: number, system?: DateSystem): Date; // UTC
```

```ts
dateToSerial(new Date(Date.UTC(2024, 0, 1))); // 45292
serialToDate(45292); // 2024-01-01T00:00:00.000Z
```

`DateSystem` is `"date1900" | "date1904"` and defaults to `date1900` (Windows Excel). Both functions accept it explicitly; the engines differ by a 1462-day offset, and the 1900 system also carries Excel's fictitious leap day. `dateToSerial` additionally accepts a duck-typed `Temporal.PlainDate` / `PlainDateTime`. The exporter uses `dateToSerial` to turn `Date` cell values into serial numbers; the preview uses `serialToDate` on the display side.

## Sheet helpers

```ts
import { sheetAddAoa } from "@marcusok/xlsx-core";

function sheetAddAoa(
  ws: Worksheet,
  data: unknown[][],
  opts?: SheetAddAoaOptions, // { origin?: string }
): void;
```

```ts
const ws = new Workbook().addSheet("Data");
sheetAddAoa(ws, [
  ["Region", "Amount"],
  ["EMEA", 1200],
]); // starts at "A1" by default
sheetAddAoa(ws, [["APAC", 800]], { origin: "A4" });
```

Without `origin`, rows append after the existing content. This is the exporter's builder primitive: it hands the engine an array-of-arrays per sheet instead of writing cells one by one.

## Typed errors

```ts
import {
  ModernXlsxError,
  LEGACY_FORMAT,
  UNRECOGNIZED_FORMAT,
  PASSWORD_PROTECTED,
  WASM_INIT_FAILED,
} from "@marcusok/xlsx-core";

class ModernXlsxError extends Error {
  readonly code: string;
  constructor(code: string, message: string);
  static fromWasmError(err: unknown): ModernXlsxError;
}

LEGACY_FORMAT; // "LEGACY_FORMAT" — legacy .xls (BIFF8), not supported
UNRECOGNIZED_FORMAT; // "UNRECOGNIZED_FORMAT" — neither ZIP nor OLE2
PASSWORD_PROTECTED; // "PASSWORD_PROTECTED" — encrypted workbook
WASM_INIT_FAILED; // "WASM_INIT_FAILED" — WASM not initialized / failed to initialize
```

Errors from the Rust core arrive in a `"[CODE] message"` form; `ModernXlsxError.fromWasmError` parses that into a `code` plus a clean message and falls back to `WASM_ERROR` when the message carries no recognizable code. The four constants are the codes exported as **values**; the engine's error-code union is larger (`ZIP_READ`, `ZIP_ENTRY`, `MISSING_PART`, `XML_PARSE`, `INVALID_*`, …) but those are only types inside the engine, so consumers that need them match on the literal string — which is what `@marcusok/excel-preview`'s error normalizer does when it maps structural failures to its own `CORRUPT` code ([preview error codes](/packages/excel-preview/api/01-create-preview#previewerror-codes)).

Always branch on `code`, not on the message: messages vary with the input shape, codes are the stable contract.

## Type re-exports

Bundled with the values above, for consumers that want to type their own code without importing `modern-xlsx`:

| Group            | Types                                                                                                                                           |
| ---------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| Read / write     | `ReadOptions`, `WriteOptions`, `WorksheetData`                                                                                                  |
| Cells and sheets | `Worksheet`, `RowData`, `CellData`, `CellType`, `ColumnInfo`, `HyperlinkData`, `SheetViewData`, `FrozenPane`                                    |
| Styles           | `StylesData`, `CellXfData`, `FontData`, `FillData`, `BorderData`, `BorderSideData`, `BorderStyle`, `AlignmentData`, `ThemeColorsData`, `NumFmt` |
| Formatting       | `FormatCellOptions`, `FormatCellResult`                                                                                                         |
| Dates            | `DateSystem`                                                                                                                                    |

```ts
import type { CellData, CellType, Worksheet } from "@marcusok/xlsx-core";
```

## Why re-export instead of depending on modern-xlsx

- **One version, one place.** `modern-xlsx` is pinned at an exact version inside this package. Re-publishing its API means the engine version is a decision made once, not one every consumer re-declares and upgrades independently.
- **One engine instance.** The engine's runtime lives in this package's `dist/`, and consumers keep the core `external`, so the identifier `readBuffer` in two different @marcusok packages resolves to the same module — and the module-global WASM initialization state is shared with the loader's.
- **No upstream `engines` conflict.** `modern-xlsx` declares `engines.node >= 24`; because its runtime is bundled in (and the declaration is only reachable through the pinned dependency of this package, whose own `engines` is `>= 22`), consumers are not pushed onto a Node version they do not run.

The trade-off is that this surface is a curated subset: symbols the engine exports but the repo does not use are not re-published, so check this page before reaching for something that exists upstream.
