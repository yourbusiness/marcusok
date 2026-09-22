import type {
  CellStyle,
  ColumnConfig,
  ExportOptions,
  ExportResult,
  ExportMode,
} from "./types";
import { WorkbookBuilder } from "./workbook-builder";
import { exportAsStream } from "./streaming-builder";
import { exportInWorker } from "./worker-exporter";
import { triggerDownload, toBlobPart } from "./download";
import { getWasmLoader } from "./wasm-loader";
import { tableExportToOptions, type TableExportOptions } from "./table-export";
import {
  echartsExportToOptions,
  type EChartsExportOptions,
} from "./echarts-export";
import { validateSheetName, validateMerges } from "./format-utils";
import { columnLabel, flattenColumnTree } from "./column-tree";
import { applyIndexColumn, assertNoReservedIndexProp } from "./sheet-normalize";

export * from "./types";
export * from "./style-presets";
export * from "./format-utils";
export * from "./table-export";
export * from "./echarts-export";
export { configureWasm, getWasmLoader } from "./wasm-loader";
export type { LoaderOptions, LoadState } from "./wasm-loader";
export { WorkbookBuilder } from "./workbook-builder";
export { exportAsStream } from "./streaming-builder";
// 序号列的展开工具（exportExcel 入口已自动调用一次）。底层入口
// WorkbookBuilder.addSheet / exportAsStream 只识别展开后的 __index__ 列，不会
// 自己展开 indexColumn；直连它们时须先 applyIndexColumn(sheet)，否则 indexColumn
// 被静默忽略。INDEX_PROP 同时导出，便于调用方避开该保留字。
export { applyIndexColumn, INDEX_PROP } from "./sheet-normalize";
// 库级基底样式：铺在所有单元格样式之下的默认对齐（见 style-utils 注释）
export { BaseCellStyle } from "./style-utils";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const STREAM_THRESHOLD = 50_000; // Workbook.toBuffer cliff starts ~55k rows
const WORKER_THRESHOLD = 20_000; // main-mode sync work is acceptable below this

/**
 * Fire-and-forget browser download. The export itself has already succeeded
 * at this point, so a trigger failure (e.g. a sandboxed DOM throwing on
 * a.click()) must not push the caller into the degradation chain and rebuild
 * the whole workbook — the Blob is already in the result. The "download"
 * phase is reported either way.
 */
function triggerDownloadIsolated(options: ExportOptions, blob: Blob): void {
  const downloadStart = performance.now();
  try {
    triggerDownload(blob, options.filename);
  } catch (e) {
    console.warn(
      `[excel-exporter] download trigger failed; the Blob is still returned in the result. Reason: ${(e as Error).message}`,
    );
  } finally {
    options.onPhase?.("download", performance.now() - downloadStart);
  }
}

type PickedMode = { mode: ExportMode; workerMode?: "workbook" | "stream" };

/**
 * Auto mode selection (verified against independent-process benchmarks).
 * - main fully blocks the thread; only for Node/SSR or browser <20,000 rows.
 * - browser >=20,000 rows go to a worker (main thread does one structured clone).
 * - inside the worker, >=50k rows use stream (avoids the toBuffer cliff).
 */
function pickMode(options: ExportOptions, totalRows: number): PickedMode {
  const explicit = options.mode ?? "auto";
  if (explicit === "stream") return { mode: "stream", workerMode: "stream" };
  if (explicit === "worker") {
    // Worker mode requires a Web Worker global. In environments without one
    // (Node/SSR), fall back to the main-thread path so styles are preserved
    // instead of silently degrading to the style-less stream fallback.
    const isBrowser =
      typeof Worker !== "undefined" && typeof window !== "undefined";
    if (!isBrowser) {
      return totalRows >= STREAM_THRESHOLD
        ? { mode: "stream", workerMode: "stream" }
        : { mode: "main" };
    }
    return {
      mode: "worker",
      workerMode: totalRows >= STREAM_THRESHOLD ? "stream" : "workbook",
    };
  }
  if (explicit === "main") return { mode: "main" };

  // auto
  const isBrowser =
    typeof Worker !== "undefined" && typeof window !== "undefined";
  if (!isBrowser) {
    return totalRows >= STREAM_THRESHOLD
      ? { mode: "stream", workerMode: "stream" }
      : { mode: "main" };
  }
  if (totalRows < WORKER_THRESHOLD) return { mode: "main" };
  if (totalRows >= STREAM_THRESHOLD)
    return { mode: "worker", workerMode: "stream" };
  return { mode: "worker", workerMode: "workbook" };
}

/**
 * 样式数值字段的前置校验（与 freezeRows / width 同类）：非法值（如
 * textRotation: 270）直通引擎会以晦涩的 serde 错误失败，并把整份导出降级
 * 为无样式 stream——正是前置校验要消除的"同一错误一边降级一边成功"分裂。
 * 范围只限 types.ts 明确承诺了取值区间的字段（textRotation 注释 0-180），
 * 不替引擎承诺未文档化的上限（如 font.size 的 409）。
 */
function validateCellStyle(style: CellStyle | undefined, at: string): void {
  if (!style) return;
  const rotation = style.alignment?.textRotation;
  if (
    rotation !== undefined &&
    (!Number.isInteger(rotation) || rotation < 0 || rotation > 180)
  ) {
    throw new Error(
      `[excel-exporter] ${at}: alignment.textRotation must be an integer between 0 and 180`,
    );
  }
  const size = style.font?.size;
  if (
    size !== undefined &&
    (typeof size !== "number" || !Number.isFinite(size) || size <= 0)
  ) {
    throw new Error(
      `[excel-exporter] ${at}: font.size must be a finite positive number`,
    );
  }
}

/** 递归校验列树（含分组列）上的样式；分组列的 style 虽被忽略，非法值一并拦截。 */
function validateColumnStyles(
  columns: ColumnConfig[],
  sheetName: string,
): void {
  for (const col of columns) {
    const at = `sheet "${sheetName}" column "${columnLabel(col)}"`;
    validateCellStyle(col.style, `${at} style`);
    validateCellStyle(col.headerStyle, `${at} headerStyle`);
    if (col.children) validateColumnStyles(col.children, sheetName);
  }
}

/**
 * Pre-flight validation of user input. Runs the same checks as the
 * Workbook/stream build paths (same functions, same messages), hoisted
 * to the entry so a configuration error fails immediately with `{ success:
 * false, error }` instead of first degrading to a stream fallback attempt
 * that re-runs the identical checks and fails identically. Engine failures
 * (WASM unavailable, build errors) still degrade to the stream as before.
 */
function validateInput(options: ExportOptions): void {
  // Guard the other core input alongside the sheets checks below: without it,
  // a JS caller omitting `filename` only fails inside triggerDownload with a
  // masked TypeError (caught as a cryptic warning), leaving success:true and
  // no file on disk. Fail fast with the structured { success: false } instead.
  if (typeof options.filename !== "string" || options.filename.length === 0) {
    throw new Error("[excel-exporter] filename must be a non-empty string");
  }
  // An empty sheets array is not a build error on the stream path (fast-xlsx
  // would zip a zero-sheet workbook Excel flags as corrupt while reporting
  // success), so reject it here like every other structural input error.
  // The same guard covers a missing/non-array `sheets`, which the caller's
  // totalRows reduce could otherwise hit first with a raw TypeError.
  if (!Array.isArray(options.sheets) || options.sheets.length === 0) {
    throw new Error("[excel-exporter] at least one sheet is required");
  }
  const seen = new Set<string>();
  for (const sheet of options.sheets) {
    // A null/primitive entry would fail on `sheet.name` with a raw TypeError;
    // give it the same clear treatment as the other structural guards below.
    if (sheet === null || typeof sheet !== "object") {
      throw new Error("[excel-exporter] each sheet must be an object");
    }
    validateSheetName(sheet.name);
    if (seen.has(sheet.name)) {
      throw new Error(`[excel-exporter] duplicate sheet name "${sheet.name}"`);
    }
    seen.add(sheet.name);
    // Guard the two arrays the build paths index into; without these, a sheet
    // missing `columns`/`data` fails downstream with a raw TypeError instead
    // of a clear, actionable message.
    if (!Array.isArray(sheet.columns)) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" must have a columns array`,
      );
    }
    if (!Array.isArray(sheet.data)) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" must have a data array`,
      );
    }
    const { leaves } = flattenColumnTree(sheet.columns);
    // 保留列在这里**无条件**拦截（与是否启用 indexColumn 无关）：构建器对
    // INDEX_PROP 的特判是无条件的，放行只会让用户自带的该列被行号静默顶替
    // （数据丢失而 success 仍为 true）。原先该检查只在 applyIndexColumn 内，
    // 未启用 indexColumn 时它直接返回原表，静默路径正好漏在最外层入口上。
    assertNoReservedIndexProp(leaves);
    validateMerges(sheet, leaves.length);
    // 样式数值字段（textRotation / font.size）前置校验，覆盖 sheet 级、列级
    //（含分组列）与 indexColumn 注入列的全部出现点。
    validateCellStyle(sheet.headerStyle, `sheet "${sheet.name}" headerStyle`);
    validateCellStyle(sheet.dataStyle, `sheet "${sheet.name}" dataStyle`);
    validateColumnStyles(sheet.columns, sheet.name);
    if (typeof sheet.indexColumn === "object" && sheet.indexColumn !== null) {
      validateCellStyle(
        sheet.indexColumn.style,
        `sheet "${sheet.name}" indexColumn style`,
      );
      validateCellStyle(
        sheet.indexColumn.headerStyle,
        `sheet "${sheet.name}" indexColumn headerStyle`,
      );
    }
    // 数值型布局/格式字段的前置校验：缺了这层，非法 width/freezeRows 会直通
    // 引擎并以晦涩的 serde 错误失败（"JSON parse error: invalid type: null,
    // expected f64"），随后整份导出被降级为无样式 stream；而同一输入在
    // >=50k 的 stream 路由却被静默忽略（width）——同一错误跨阈值一边降级
    // 一边成功。在此拦截使两条路径行为一致，且报错能定位到具体的列/表。
    if (
      sheet.freezeRows !== undefined &&
      (typeof sheet.freezeRows !== "number" ||
        !Number.isInteger(sheet.freezeRows) ||
        sheet.freezeRows < 0)
    ) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" freezeRows must be a non-negative integer`,
      );
    }
    // indexColumn 形状校验：非 boolean 非（非 null）object 的值（JS 调用方
    // 可能传 null / 字符串 / 数字）若放行，会被归一化与构建器各自按不同
    // 方式静默吞掉——与 freezeRows 等字段一致，前置报错。
    if (
      sheet.indexColumn !== undefined &&
      typeof sheet.indexColumn !== "boolean" &&
      (typeof sheet.indexColumn !== "object" || sheet.indexColumn === null)
    ) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" indexColumn must be a boolean or an options object`,
      );
    }
    // indexColumn.start 与 freezeRows 同类数值校验：非整数/负数只会在构建期
    // 产出 1.5、-1 这类怪序号，前置拦截。
    if (
      typeof sheet.indexColumn === "object" &&
      sheet.indexColumn !== null &&
      sheet.indexColumn.start !== undefined &&
      (!Number.isInteger(sheet.indexColumn.start) ||
        sheet.indexColumn.start < 0)
    ) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" indexColumn.start must be a non-negative integer`,
      );
    }
    // indexColumn.width 与下方 col.width 同类，但序号列是在**本校验之后**才由
    // applyIndexColumn 注入的（见 exportExcel），不在这里单独覆盖就会绕过校验：
    // NaN 直通引擎以 serde 错误失败并把整份导出降级为无样式 stream，而 >=50k 的
    // stream 路由又把 width 当作被丢弃的特性静默忽略——同一输入跨阈值一边降级
    // 一边成功，正是 col.width 校验要消除的那种分裂。
    if (
      typeof sheet.indexColumn === "object" &&
      sheet.indexColumn !== null &&
      sheet.indexColumn.width !== undefined &&
      (typeof sheet.indexColumn.width !== "number" ||
        !Number.isFinite(sheet.indexColumn.width) ||
        sheet.indexColumn.width < 0)
    ) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" indexColumn.width must be a finite non-negative number`,
      );
    }
    for (const col of leaves) {
      // OOXML 中 width=0 合法（隐藏列），负数与非有限数非法；只有叶子列
      // 消费 width，分组列上的 width 本就被忽略。
      if (
        col.width !== undefined &&
        (typeof col.width !== "number" ||
          !Number.isFinite(col.width) ||
          col.width < 0)
      ) {
        throw new Error(
          `[excel-exporter] column "${columnLabel(col)}" width must be a finite non-negative number`,
        );
      }
      if (col.format && typeof col.format === "object") {
        const spec = col.format;
        // decimals 两路径共享：Workbook 路径拼 numFormat 字符串（任意值都
        // "能出"），stream 路径烧入 toFixed(decimals)——收敛到 toFixed 自身
        // 的 0..100 上限，同一 spec 才不会在 50k 行上下一边成功一边抛错。
        if (
          spec.type === "number" &&
          spec.decimals !== undefined &&
          (!Number.isInteger(spec.decimals) ||
            spec.decimals < 0 ||
            spec.decimals > 100)
        ) {
          throw new Error(
            `[excel-exporter] column "${columnLabel(col)}" format.decimals must be an integer between 0 and 100`,
          );
        }
        // padding.length：非整数/负数会让 padStart 抛 RangeError 或静默不
        // 填充，超大值会生成巨型字符串，均在渲染期才爆——前置拦截。
        if (
          spec.type === "padding" &&
          (!Number.isInteger(spec.length) ||
            spec.length < 0 ||
            spec.length > 10_000)
        ) {
          throw new Error(
            `[excel-exporter] column "${columnLabel(col)}" format.length must be an integer between 0 and 10000`,
          );
        }
      }
    }
  }
}

/**
 * Export to Excel (main entry).
 *
 * @example
 * ```ts
 * import { exportExcel, StylePresets } from '@marcusok/excel-exporter';
 *
 * await exportExcel({
 *   filename: 'sales-report',
 *   sheets: [{
 *     name: 'Sales', freezeRows: 1, autoFilter: true,
 *     columns: [
 *       { prop: 'product', label: 'Product', width: 20 },
 *       { prop: 'revenue', label: 'Revenue', width: 15, style: StylePresets.currency },
 *     ],
 *     data: [{ product: 'Widget', revenue: 9999.99 }],
 *   }],
 * });
 * ```
 */
export async function exportExcel(
  options: ExportOptions,
): Promise<ExportResult> {
  const start = performance.now();

  // Leading 0 fires exactly once here, on every route (the stream fallback
  // included), so consumers always see the documented 0 -> ... -> 1 pair.
  options.onProgress?.(0);

  // Invalid input fails here on every route (same messages as before; the
  // build paths keep their own checks for direct callers). The trailing 1 is
  // still emitted so the 0 -> 1 progress contract holds for failed exports.
  //
  // Validation runs BEFORE totalRows is computed: reduce()ing a non-array
  // `sheets` (or a sheet without a `data` array) would throw a raw TypeError
  // that rejects the promise, bypassing the structured { success: false }
  // contract — validateInput guards those shapes first, so the computation
  // below only sees well-formed input.
  try {
    validateInput(options);
    // 序号列归一化紧跟校验：INDEX_PROP 冲突在此以结构化 { success: false }
    // 失败（applyIndexColumn 抛错）。只在此处展开一次，worker / stream 路由
    // 拿到的 sheets 已含虚拟列，构建器只需特判 INDEX_PROP 取行号。
    options = { ...options, sheets: options.sheets.map(applyIndexColumn) };
  } catch (e) {
    options.onProgress?.(1);
    return {
      success: false,
      error: e as Error,
      duration: performance.now() - start,
    };
  }

  const totalRows = options.sheets.reduce((s, sh) => s + sh.data.length, 0);

  const picked = pickMode(options, totalRows);
  const needsWasm = picked.workerMode !== "stream";
  const loader = getWasmLoader();

  /**
   * Execute the export on this thread (Workbook build, or the WASM-free fast
   * stream). Used by the Node/SSR route, the browser main route, and as the
   * style-preserving retry when the browser worker path fails; `forceStream`
   * selects the pure-JS stream for the terminal degradation (finishWithStream).
   * Throws on failure; callers decide the next degradation step.
   */
  const runOnMainThread = async (
    forceStream = false,
  ): Promise<ExportResult> => {
    const useStream = forceStream || picked.workerMode === "stream";
    if (!useStream) {
      const initStart = performance.now();
      await loader.ensureLoaded();
      options.onPhase?.("init", performance.now() - initStart);
    } else {
      // Fast stream does not use WASM; report an empty init phase so the
      // public phase sequence remains stable across main/stream routes.
      options.onPhase?.("init", 0);
    }
    let result: ExportResult;
    const buildStart = performance.now();
    try {
      if (useStream) {
        const { bytes, rowCount } = await exportAsStream(
          options.sheets,
          options.onProgress,
        );
        result = {
          success: true,
          blob: new Blob([toBlobPart(bytes)], { type: XLSX_MIME }),
          engine: "modern-xlsx",
          mode: "stream",
          duration: performance.now() - start,
          rowCount,
        };
      } else {
        const builder = await WorkbookBuilder.create();
        options.sheets.forEach((s) => builder.addSheet(s));
        const bytes = await builder.toBuffer();
        result = {
          success: true,
          blob: new Blob([toBlobPart(bytes)], { type: XLSX_MIME }),
          engine: "modern-xlsx",
          mode: "main",
          duration: performance.now() - start,
          rowCount: totalRows,
        };
      }
    } finally {
      // Reported even when the build throws, so a failed attempt that falls
      // back to the stream still shows how long it spent before failing.
      options.onPhase?.("build", performance.now() - buildStart);
    }
    options.onProgress?.(1);
    // Node has no document: triggerDownload would be a no-op, so neither the
    // click nor the "download" phase is reported (matches ExportPhase docs).
    if (options.download !== false && typeof document !== "undefined") {
      triggerDownloadIsolated(options, result.blob!);
    }
    return result;
  };

  // Terminal degradation: the pure-JS fast stream on the main thread. It
  // covers the old SheetJS fallback's surface (headers/merges preserved,
  // styles and layout features dropped) with none of its costs — no optional
  // peer dependency, no runtime CDN load, no network access at all. Like the
  // SheetJS path before it, it never leaves the trailing progress 1 dangling:
  // runOnMainThread emits it on success, the catch below on failure.
  // (Declared after runOnMainThread — it closes over it.)
  const finishWithStream = async (reason: string): Promise<ExportResult> => {
    console.warn(
      `[excel-exporter] Falling back to the style-less fast stream. Reason: ${reason}`,
    );
    try {
      const result = await runOnMainThread(true);
      return {
        ...result,
        // Surface the degradation programmatically (parity with the old
        // SheetJS fallback's soft error): success stays true.
        error: new Error(
          `Fallback: styles stripped (fast stream). Reason: ${reason}`,
        ),
      };
    } catch (e) {
      options.onProgress?.(1);
      return {
        success: false,
        error: e as Error,
        duration: performance.now() - start,
      };
    }
  };

  // WASM unsupported and the chosen route needs it: degrade straight to the
  // WASM-free stream (checked after the closures above are initialized).
  if (needsWasm && !loader.supported) {
    return finishWithStream("WebAssembly not supported");
  }

  /**
   * Worker-path degradation chain: retry on the main thread first (modern-xlsx
   * keeps styles), and only when that also fails degrade to the style-less
   * fast stream. The trailing progress 1 is emitted exactly once on either
   * sub-route (runOnMainThread on success, or finishWithStream's catch).
   */
  const retryOnMainThread = async (reason: string): Promise<ExportResult> => {
    console.warn(
      `[excel-exporter] Worker path failed (${reason}); retrying on the main thread to preserve styles`,
    );
    try {
      return await runOnMainThread();
    } catch (e) {
      // When the worker was on the stream route, the main-thread retry just ran
      // the very same fast stream on the very same input: a third attempt via
      // finishWithStream would fail deterministically, so fail here instead of
      // paying for (and logging) a doomed extra build.
      if (picked.workerMode === "stream") {
        options.onProgress?.(1);
        return {
          success: false,
          error: new Error(
            `${reason}; main-thread stream retry failed: ${(e as Error).message}`,
          ),
          duration: performance.now() - start,
        };
      }
      return finishWithStream(
        `${reason}; main-thread retry failed: ${(e as Error).message}`,
      );
    }
  };

  // Node main/stream: execute directly on this thread (no Web Worker available).
  if (
    picked.mode === "main" ||
    (picked.mode === "stream" && typeof window === "undefined")
  ) {
    try {
      return await runOnMainThread();
    } catch (e) {
      // 与下方 retryOnMainThread 的防护对齐：首次尝试已经是 fast stream
      // （Node stream 路由）时，finishWithStream 重试等于对同一输入重跑一遍
      // 确定性失败的构建——直接失败，不再白付第二次构建的时间与告警。
      if (picked.workerMode === "stream") {
        options.onProgress?.(1);
        return {
          success: false,
          error: e as Error,
          duration: performance.now() - start,
        };
      }
      return finishWithStream((e as Error).message);
    }
  }

  // Browser worker mode: offload to worker (main thread does one structured clone).
  try {
    const result = await exportInWorker(options, picked.workerMode!);
    if (result.success) {
      // Terminal 1 on success only: the failure route hands the sequence to
      // the degradation chain, which emits it exactly once (the types.ts
      // contract) -- emitting it here too duplicated the trailing 1.
      options.onProgress?.(1);
      if (options.download !== false) {
        triggerDownloadIsolated(options, result.blob!);
      }
      return result;
    }
    // Worker export failed (e.g. WASM init error inside the worker) -> retry on
    // the main thread before degrading to the fast stream.
    return retryOnMainThread(result.error?.message ?? "worker export failed");
  } catch (e) {
    return retryOnMainThread((e as Error).message);
  }
}

/**
 * Convenience wrapper for common table data shapes.
 *
 * Accepts Ant Design-style (`title` / `dataIndex`) and Element Plus-style
 * (`label` / `prop`) column descriptors, normalizes them to `SheetConfig`,
 * and delegates to {@link exportExcel}.
 */
export async function exportTable(
  options: TableExportOptions,
): Promise<ExportResult> {
  return exportExcel(tableExportToOptions(options));
}

/**
 * Convenience wrapper for a small, explicit subset of ECharts options.
 *
 * Supports category-axis series data, pie-like name/value data, and
 * scatter-like coordinate pairs. Unsupported `dataset` mode throws instead of
 * guessing.
 */
export async function exportEcharts(
  options: EChartsExportOptions,
): Promise<ExportResult> {
  return exportExcel(echartsExportToOptions(options));
}
