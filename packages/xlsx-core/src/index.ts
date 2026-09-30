/**
 * @marcusok/xlsx-core — the repo's single modern-xlsx integration point.
 *
 * The engine (modern-xlsx, Rust + WASM) is bundled INTO this package's dist
 * at build time, and this package is a PRIVATE workspace layer (not on npm):
 * the two business packages bundle its dist into their own builds, so
 * consumers depend on @marcusok/excel-exporter / @marcusok/excel-preview
 * alone and are immune to modern-xlsx's engines.node>=24 declaration (the
 * business packages redeclare >=22). The pinned `modern-xlsx` dependency
 * stays declared in the business packages' package.json solely for
 * consumer-side d.ts resolution (tsup leaves type re-exports as external
 * imports — see those packages' READMEs for the full rationale), not for
 * runtime.
 *
 * Trade-off of bundling-in (instead of keeping this package external):
 * a page using BOTH business packages carries two engine instances and two
 * WASM binaries on the main thread. Most admin pages use only one of the
 * two, and content-hash assets plus the HTTP cache usually dedupe the
 * binary transfer — accepted to keep the npm release surface at two
 * packages with no cross-package version coupling. Worker entrypoints have
 * always carried their own copy anyway (browser module workers cannot
 * resolve bare specifiers).
 *
 * Everything below is a re-export surface: modern-xlsx's public API (the
 * subset shared by the repo's packages, plus their type needs) and the WASM
 * loader. No behavior of its own beyond the loader.
 */

// ---------------------------------------------------------------------------
// WASM loader (browser fetch / Node sync-init dual path, timeout + retry)
// ---------------------------------------------------------------------------
export {
  WasmLoader,
  configureWasm,
  getWasmLoader,
  defaultWasmUrl,
} from "./wasm-loader";
export type { LoaderOptions, LoadState } from "./wasm-loader";

// ---------------------------------------------------------------------------
// modern-xlsx engine re-exports (runtime values)
// ---------------------------------------------------------------------------
export {
  Workbook,
  readBuffer,
  // wasm init — consumed by worker entrypoints and the loader's tests
  initWasm,
  initWasmSync,
  // cell / range reference utilities
  encodeCellRef,
  decodeCellRef,
  encodeRange,
  decodeRange,
  columnToLetter,
  letterToColumn,
  // formatting engine (preview's numFmt layer + exporter's format utils)
  formatCell,
  formatCellRich,
  getBuiltinFormat,
  loadFormatTable,
  isDateFormatId,
  isDateFormatCode,
  // date serial conversion (epoch handling, see serialToDate notes)
  serialToDate,
  dateToSerial,
  // sheet helpers used by the exporter's builder
  sheetAddAoa,
  // typed error surface: preview maps err.code to friendly messages
  ModernXlsxError,
} from "modern-xlsx";

// engine error-code constants (string values exported by modern-xlsx)
export {
  LEGACY_FORMAT,
  UNRECOGNIZED_FORMAT,
  PASSWORD_PROTECTED,
  WASM_INIT_FAILED,
} from "modern-xlsx";

// ---------------------------------------------------------------------------
// modern-xlsx type re-exports
// ---------------------------------------------------------------------------
export type {
  Worksheet,
  BorderStyle,
  ReadOptions,
  WriteOptions,
  WorksheetData,
  RowData,
  CellData,
  ColumnInfo,
  StylesData,
  CellXfData,
  FontData,
  FillData,
  BorderData,
  BorderSideData,
  AlignmentData,
  ThemeColorsData,
  DateSystem,
  CellType,
  HyperlinkData,
  SheetViewData,
  FrozenPane,
  NumFmt,
  FormatCellResult,
  FormatCellOptions,
} from "modern-xlsx";
