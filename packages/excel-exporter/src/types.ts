/**
 * Type definitions for @marcusok/excel-exporter.
 *
 * Colors use 6-digit RGB hex (e.g. `'FF0000'`), matching modern-xlsx's
 * FontData.color / FillData.fgColor / BorderSideData.color (verified from
 * dist/validate-chart-D1O7LOfU.d.mts @ modern-xlsx 1.2.0).
 */

/**
 * Cell border line style. Inlined from modern-xlsx's BorderSideData
 * (dist/validate-chart-D1O7LOfU.d.mts @ 1.2.0) so the published `.d.ts` has
 * no dependency imports — the engine itself is bundled into this package.
 */
export type BorderStyle =
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

/** Business-friendly cell style config; mapped to StyleBuilder in style-utils.ts. */
export interface CellStyle {
  font?: {
    bold?: boolean;
    italic?: boolean;
    size?: number;
    color?: string; // 6-digit RGB hex, e.g. 'FF0000'
    name?: string; // font name, e.g. 'Arial'
  };
  fill?: {
    pattern?: "solid" | "none";
    fgColor?: string; // 6-digit RGB hex
    bgColor?: string;
  };
  alignment?: {
    horizontal?: "left" | "center" | "right";
    vertical?: "top" | "center" | "bottom";
    wrapText?: boolean;
    textRotation?: number; // 0-180
  };
  border?: {
    top?: { style: BorderStyle; color?: string };
    bottom?: { style: BorderStyle; color?: string };
    left?: { style: BorderStyle; color?: string };
    right?: { style: BorderStyle; color?: string };
  };
  numFormat?: string; // e.g. '#,##0.00', 'yyyy-mm-dd', '0.00%'
}

/**
 * Worker-compatible, data-describing format spec. Functions cannot cross the
 * structured-clone boundary into a Web Worker, so routes that enter a worker
 * (browser >=20,000 rows in auto mode, or an explicit worker/stream mode in a
 * browser) strip function-form formats with a console warning — use FormatSpec
 * there. Main-thread routes execute the function form: `main` mode (browser
 * <20,000 rows / Node), the Node main-thread stream (auto >=50,000 rows), and
 * main-thread retries after a worker failure.
 *
 * Date semantics: `date`/`datetime` interpret values by their **UTC
 * components**. The workbook path serializes via modern-xlsx's `dateToSerial`
 * (UTC wall clock) and the stream path formats the same UTC components
 * into strings, so all paths agree in every timezone. Date-only ISO strings
 * ("2025-01-05") parse as UTC midnight per ECMA-262; prefer them (or
 * `Date.UTC(...)`) over locally-constructed Dates, whose UTC components can
 * fall on the previous day in non-UTC timezones.
 *
 * Pattern tokens: the stream path (>= 50,000 rows, explicit stream mode, or
 * the fallback) lower-cases the whole pattern first, then parses only
 * `yyyy`/`MM`/`dd`/`HH`/`mm`/`ss` (`mm` resolves to minutes vs month by
 * context; `yy` and the single-letter `M`/`d` are NOT tokens) and emits
 * everything else as-is — lower-cased, and with quoted literals NOT
 * interpreted — so `yyyy"年"M"月"d"日"` renders as `2026"年"m"月"d"日"` while
 * the Workbook path hands the same pattern to Excel as a numFormat and renders
 * `2026年7月1日`. Superset tokens are likewise only partially passed through:
 * `mmm` parses its `mm` prefix and emits a stray `m` (`"mmm"` -> `"09m"`),
 * where Excel renders the month abbreviation. Stick to the six tokens for
 * cross-threshold consistency.
 */
export type FormatSpec =
  | { type: "enum"; map: Record<string, string>; fallback?: string }
  | { type: "date"; pattern?: string } // default 'yyyy-MM-dd'
  | { type: "datetime"; pattern?: string } // default 'yyyy-MM-dd HH:mm'
  | { type: "number"; decimals?: number; thousands?: boolean }
  | { type: "padding"; fill: string; length: number; align?: "left" | "right" };

/** Column configuration. A column with `children` is a group header; leaf columns produce data cells. */
export interface ColumnConfig {
  /**
   * Data row field name (Element Plus naming). Required for leaf columns
   * (validated at export time); group columns (with `children`) may omit it.
   */
  prop?: string;
  /**
   * Legacy alias of `prop` (pre-2.2 naming), kept for backward compatibility.
   * `prop` takes precedence when both are present.
   * @deprecated use `prop`
   */
  key?: string;
  /**
   * Header text (leaf or group), Element Plus naming. At least one of
   * `label` / legacy `header` must be provided (validated at export time).
   */
  label?: string;
  /**
   * Legacy alias of `label` (pre-2.2 naming), kept for backward compatibility.
   * `label` takes precedence when both are present.
   * @deprecated use `label`
   */
  header?: string;
  /**
   * Group header: the column tree becomes multi-row headers, and each group
   * header cell is merged across its descendant leaf columns. `children: []`
   * is treated as a leaf column.
   */
  children?: ColumnConfig[];
  /** Column width in Excel character units. Leaf columns only. Mapped to ws.setColumnWidth(col, width) (1-based). */
  width?: number;
  /** Style applied to all data cells in this column (not the header). Leaf columns only. */
  style?: CellStyle;
  /**
   * Style applied to this column's header cell(s) (group header cells
   * included). **Replaces** `SheetConfig.headerStyle` wholesale when present —
   * unlike `style`, which deep-merges over `dataStyle` field by field (see
   * style-utils.ts: deliberately asymmetric).
   */
  headerStyle?: CellStyle;
  /**
   * Value formatter: FormatSpec (worker-compatible) or function (main-thread
   * routes only; stripped with a warning on worker routes — see FormatSpec).
   *
   * Cross-path precision: a `{ type: "number" }` spec without `decimals`
   * defaults to 0, but only the stream path (>= STREAM_THRESHOLD, 50,000 rows)
   * bakes `toFixed(0)` into the stored cell value. The Workbook path keeps full
   * precision and renders decimals via numFormat, so the same spec can store
   * `9999.99` (Workbook) vs `10000` (stream). Always set `decimals` explicitly
   * for cross-threshold consistency (see docs/excel-export-design.md 4.8).
   *
   * Cross-path `thousands`: the Workbook path renders the separator via an
   * auto-injected `#,##0` numFormat; the stream path (>=50k rows /
   * degraded exports) cannot use numFormat and keeps the cell a *number*, so
   * they render `9999.99` without separators. Baking separators into the value
   * would turn data cells into text and break downstream calculations, so the
   * difference is intentional — do not rely on visible separators above the
   * 50k threshold.
   *
   * Missing values in numeric-ish specs: `null`/`undefined` (and blank/whitespace-only
   * strings — the most common missing-value shape coming out of databases, forms
   * and CSV imports; `Number("") === 0` would silently turn them into a meaningful
   * `0`) render as empty cells on every path in `{ type: "number" }` columns.
   * `{ type: "padding" }` likewise leaves them empty instead of padding the empty
   * string into a fake-looking `"00000"`. `{ type: "enum" }` maps them through the
   * `""` key like any other value (hit your `fallback` unless `""` is mapped).
   */
  format?:
    | FormatSpec
    | ((
        value: unknown,
        row: Record<string, unknown>,
      ) => string | number | boolean);
}

/** Merge range: relative to the data area, row/col are 0-based (row 0 = first data row). */
export interface MergeRange {
  row: number;
  col: number;
  rowspan: number;
  colspan: number;
}

/**
 * Options for the auto-injected leading row-number column
 * (`SheetConfig.indexColumn`). The shorthand `true` equals `{}` — every
 * option has a default.
 */
export interface IndexColumnOptions {
  /** Header text; defaults to `"序号"`. */
  label?: string;
  /** Column width in Excel character units; defaults to `6`. */
  width?: number;
  /**
   * Number shown on the first data row; row i displays `start + i`.
   * Defaults to `1`; must be a non-negative integer (validated at export time).
   */
  start?: number;
  /** Style for the index column's data cells; merged over `SheetConfig.dataStyle` like any column style. */
  style?: CellStyle;
  /** Style for the index column's header cell; overrides `SheetConfig.headerStyle`. */
  headerStyle?: CellStyle;
}

/** Sheet configuration. */
export interface SheetConfig {
  name: string; // 1-31 chars, ECMA-376 validation (no `: \ / ? * [ ]`, no leading/trailing apostrophe)
  columns: ColumnConfig[];
  /**
   * Data rows keyed by column `prop`. Cell values are normalized identically on
   * every export path (main / worker / stream, including the stream fallback):
   * non-finite numbers (NaN/Infinity), plain objects, `Date`s and bigints
   * without a `format` are written as their visible string form (JSON for
   * objects, ISO for Dates), so a dataset crossing the 50k-row threshold keeps
   * the same content.
   */
  data: Record<string, unknown>[];
  /**
   * Style applied to every header cell. A column-level `headerStyle` replaces
   * this wholesale when present (no field-level merge — see ColumnConfig.headerStyle).
   */
  headerStyle?: CellStyle;
  /**
   * Base style applied to every data cell; a column's own `style` is
   * deep-merged over it field by field (so a table-wide border survives a
   * column that only sets `numFormat`, and vice versa). Header cells are not
   * affected — use `headerStyle`. Dropped with a warning on the stream path.
   */
  dataStyle?: CellStyle;
  /**
   * Inject a leading row-number column (1..N) without touching `data` or
   * `columns`: the column is inserted internally, values are generated from
   * the row number (never read from `data`), and existing `merges` are
   * shifted right by one column automatically. Works on every export path
   * (workbook / worker / stream); styles on it follow `dataStyle` /
   * `headerStyle` semantics. Pass `true` for defaults, or an object to
   * customize label / width / start / styles.
   */
  indexColumn?: boolean | IndexColumnOptions;
  /** Number of header rows to freeze (usually 1). Maps to ws.frozenPane = { rows, cols: 0 }. */
  freezeRows?: number;
  /** Merged cell ranges. */
  merges?: MergeRange[];
  /**
   * Whether to add an auto-filter. The filter range spans the last header row
   * plus all data rows (Excel's filter semantics — the dropdown sits on the
   * header row and covers the data beneath it), not the header rows alone.
   */
  autoFilter?: boolean;
}

/** Export mode. */
export type ExportMode = "auto" | "main" | "worker" | "stream";

/**
 * Named export stages, reported through `onPhase` as they complete. Phases are
 * strictly sequential within one export call.
 *
 * - `"init"`: WASM initialization. Main-thread paths measure
 *   `loader.ensureLoaded()`; worker mode measures the worker's `initWasm()`
 *   (only reported when the worker actually re-initializes, not when its WASM
 *   instance is already cached). Reported as a zero-duration phase by the
 *   WASM-free stream fallback and the main-thread stream routes. The browser
 *   worker-stream route (stream engine inside a worker) uses no WASM and
 *   reports no `"init"` at all.
 * - `"build"`: workbook construction. Covers the Workbook/stream builder.
 *   Each main-thread build attempt reports its own `"build"` phase (in a
 *   `finally`, so also when that attempt throws), so a degradation chain
 *   (e.g. failed worker build -> main-thread retry -> stream fallback) reports
 *   one phase per main-thread attempt. An attempt that fails *inside* the
 *   worker reports no `"build"` phase — its failure surfaces only through the
 *   retry's error.
 * - `"download"`: the synchronous browser download trigger
 *   (`triggerDownload`); only reported when `download !== false`. Not reported
 *   in Node (no `document`).
 */
export type ExportPhase = "init" | "build" | "download";

/** Export options. */
export interface ExportOptions {
  sheets: SheetConfig[];
  /** Download file name; `.xlsx` appended unless already present. Validated as a non-empty string at export time. */
  filename: string;
  /** Mode selection: auto = auto-decide by row count (default). */
  mode?: ExportMode;
  /**
   * Progress callback (0-1). The leading 0 and the trailing 1 are each emitted
   * exactly once by `exportExcel` itself, on every route — including the
   * stream fallback and exports that ultimately fail — so a progress UI can
   * always be closed on the final 1. The stream path additionally reports
   * intermediate values every 1,000 rows; the checkpoint that would land on
   * the final row is skipped, so the trailing 1 is never duplicated.
   */
  onProgress?: (progress: number) => void;
  /**
   * Optional per-stage timing callback. Receives the phase name and its
   * wall-clock duration in ms (0 means the phase did no work, e.g. WASM was
   * already loaded). Useful for metrics/play panels; does not affect
   * `ExportResult.duration` (which measures the whole export on main-thread
   * routes; the worker route's duration is measured on the main thread from
   * the call into `exportInWorker` — including the pre-post serialization and
   * the worker round-trip — through Blob construction, so it is wider than
   * the pure in-worker build time).
   */
  onPhase?: (phase: ExportPhase, durationMs: number) => void;
  /** Trigger browser download (default true). Set false to only return a Blob. */
  download?: boolean;
}

/** Export result. */
export interface ExportResult {
  success: boolean;
  blob?: Blob;
  /** Engine actually used. */
  engine?: "modern-xlsx";
  /** Mode actually used. */
  mode?: ExportMode;
  duration?: number; // ms
  rowCount?: number;
  /**
   * Failure cause when `success` is false. Also set — together with
   * `success: true` — when the export succeeded via the style-less stream
   * fallback, carrying the degradation reason (e.g. "WebAssembly not
   * supported") so callers can monitor the fallback rate programmatically.
   * Check `success` first; a present `error` alone does not mean the export
   * failed.
   */
  error?: Error;
}
