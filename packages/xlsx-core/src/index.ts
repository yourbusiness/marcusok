/**
 * @marcusok/xlsx-core — the repo's single modern-xlsx integration point.
 *
 * The engine (modern-xlsx, Rust + WASM) is bundled INTO this package's dist
 * at build time, so consumers depend on @marcusok/xlsx-core alone: they see
 * zero external runtime dependencies and are immune to modern-xlsx's
 * engines.node>=24 declaration. The pinned `modern-xlsx` dependency stays
 * declared in package.json solely for consumer-side d.ts resolution (tsup
 * leaves type re-exports as external imports — see the package README for
 * the full rationale), not for runtime. Packages that consume this core
 * mark it `external` in their main builds, so a page using several
 * @marcusok packages loads one engine instance (and one WASM binary) on the
 * main thread. Worker entrypoints of consuming packages must stay
 * single-file-self-contained (browser module workers cannot resolve bare
 * specifiers), so they bundle this core — the duplication there is inherent
 * and unchanged from the pre-core era.
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
