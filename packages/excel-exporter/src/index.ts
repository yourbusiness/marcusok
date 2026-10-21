import type { ExportOptions, ExportResult } from "./types";
import { WorkbookBuilder } from "./workbook-builder";
import { exportAsStream } from "./streaming-builder";
import {
  exportInWorker,
  isAbortError,
  type NoFallbackError,
} from "./worker-exporter";
import { toBlobPart } from "./download";
import {
  triggerDownloadIsolated,
  mergeOverlayText,
  overlayPhaseKey,
} from "./overlay-wiring";
import { getWasmLoader } from "@marcusok/xlsx-core";
import { tableExportToOptions, type TableExportOptions } from "./table-export";
import {
  echartsExportToOptions,
  type EChartsExportOptions,
} from "./echarts-export";
import { applyIndexColumn } from "./sheet-normalize";
import { validateInput } from "./input-validation";
import { estimateTotalCells, pickMode } from "./routing";
import {
  nextPaint,
  showProgressOverlay,
  type ProgressOverlayOptions,
} from "@marcusok/progress-overlay";

export * from "./types";
export * from "./style-presets";
export * from "./format-utils";
export * from "./table-export";
export * from "./echarts-export";
// WASM 加载与公共引擎面统一由共享底层包提供（对消费方 API 面不变：同名
// re-export，签名与行为一致，来源从本包 internals 换成 @marcusok/xlsx-core）。
export { configureWasm, getWasmLoader } from "@marcusok/xlsx-core";
export type { LoaderOptions, LoadState } from "@marcusok/xlsx-core";
// 遮罩类型随 `overlay` 选项暴露：调用方定制它时无需再从子路径或公共包取类型。
export type {
  ProgressOverlayHandle,
  ProgressOverlayOptions,
  ProgressOverlayTextOptions,
} from "@marcusok/progress-overlay";
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
/**
 * Export to Excel (main entry).
 *
 * 自 2.8.0 起默认显示全屏进度遮罩（浏览器环境；Node/SSR 无 document 时
 * 无操作）。`overlay: false` 完全关闭；传 {@link ProgressOverlayOptions}
 * 定制文案/延迟/主题。遮罩能力抽自 @marcusok/progress-overlay。
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
  // null/undefined options（JS 调用方）直接透传核心：validateInput 会以结构化
  // { success: false } 拒绝。必须在接线**之前**早退——下方 `{ ...options }`
  // 的展开会把 null 变成合法对象、改变报错口径（options must be an object
  // → filename must be a non-empty string）。
  if (options === null || typeof options !== "object") {
    return exportExcelCore(options);
  }
  // overlay: false → 零开销直通：不挂遮罩、不让帧、回调链原样。
  if (options.overlay === false) return exportExcelCore(options);
  const overlayOptions: ProgressOverlayOptions =
    options?.overlay === undefined || options?.overlay === true
      ? {}
      : options.overlay;

  // Node 无 document 时 showProgressOverlay 返回 NOOP 句柄，接线零副作用；
  // willDownload 的口径与 download 阶段一致（types.ts：download !== false
  // 且有 document 才上报）。
  const willDownload =
    options?.download !== false && typeof document !== "undefined";
  const overlay = showProgressOverlay({
    ...overlayOptions,
    text: mergeOverlayText(overlayOptions.text),
  });
  try {
    // 先让出一帧：遮罩挂载与紧接其后的同步构建若落在同一个任务里，浏览器
    // 永远不会绘制遮罩（用户只看到页面卡死）。delayMs 为 0 时这一步是遮罩
    // 能被看见的唯一保证。
    await nextPaint();
    return await exportExcelCore({
      ...options,
      // 链式追加而非替换：调用方挂在自己回调上的指标面板照常工作。
      onProgress: (progress) => {
        overlay.setProgress(progress);
        options?.onProgress?.(progress);
      },
      onPhase: (phase, durationMs) => {
        overlay.setPhase(overlayPhaseKey(phase, willDownload));
        options?.onPhase?.(phase, durationMs);
      },
    });
  } finally {
    // 关闭时机只能在这里：收尾的 onProgress(1) 在失败路径同样会发（types.ts
    // 契约），成功后 download 还排在它之后，用进度值判定关闭都会出错。
    overlay.close();
  }
}

/**
 * 用户回调（onProgress/onPhase）异常隔离。回调是调用方代码，其抛错（常见如
 * 回调驱动的 UI 组件已卸载）不得影响导出本体：首个 onProgress(0) 抛错会以
 * 裸异常 reject、突破"永远 resolve 结构化结果"的契约；更隐蔽的是 worker
 * 分支已成功拿到 Blob 后收尾的 onProgress(1) 抛错会落入 catch，把整份工作簿
 * 在主线程全量重建甚至降级剥样式。在核心入口统一包一层后，主线程 / worker
 * 消息回调 / 流式循环三条路径上的全部调用点行为一致：抛错仅 warn。
 */
function guardCallback<A extends unknown[]>(
  cb: ((...args: A) => void) | undefined,
): ((...args: A) => void) | undefined {
  if (!cb) return cb;
  return (...args: A) => {
    try {
      cb(...args);
    } catch (e) {
      console.warn("[excel-exporter] user callback threw; ignoring", e);
    }
  };
}

/** {@link exportExcel} 的路由与构建主体（overlay 接线在包装层完成）。 */
async function exportExcelCore(options: ExportOptions): Promise<ExportResult> {
  const start = performance.now();

  // 回调隔离必须先于首个 onProgress(0) 接线；null options 无回调可包，
  // 且下方可选链与 validateInput 已覆盖该形态。
  if (options !== null && typeof options === "object") {
    options = {
      ...options,
      onProgress: guardCallback(options.onProgress),
      onPhase: guardCallback(options.onPhase),
    };
  }

  // Leading 0 fires exactly once here, on every route (the stream fallback
  // included), so consumers always see the documented 0 -> ... -> 1 pair.
  // The outer optional chain also covers a null/undefined `options` from a JS
  // caller: dereferencing it here (before the try) would reject with a raw
  // TypeError, bypassing the structured { success: false } contract the
  // comment below promises; validateInput reports it properly instead.
  options?.onProgress?.(0);

  /** 用户取消：统一的结构化失败，且不再进入降级链。 */
  const aborted = (): ExportResult => {
    options?.onProgress?.(1);
    const err = new Error("[excel-exporter] export aborted by the caller");
    err.name = "AbortError";
    return { success: false, error: err, duration: performance.now() - start };
  };
  const abortedError = (): boolean =>
    options !== null &&
    typeof options === "object" &&
    options.signal?.aborted === true;
  const isAbort = (e: unknown): boolean => isAbortError(e) || abortedError();

  if (abortedError()) return aborted();

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
    // options 可能正是校验失败的原因（null/undefined），可选链同入口处的
    // 首个 0 一样兜住空 options。
    options?.onProgress?.(1);
    return {
      success: false,
      error: e as Error,
      duration: performance.now() - start,
    };
  }

  const totalRows = options.sheets.reduce((s, sh) => s + sh.data.length, 0);
  // 路由同时看行数与单元格总量（宽表会被行数单指标误送进 Workbook 重路径），
  // 详见 routing.ts 的阈值依据。applyIndexColumn 已展开，叶子列数即终值。
  const totalCells = estimateTotalCells(options.sheets);

  const picked = pickMode(options.mode, totalRows, totalCells);
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
    if (abortedError()) throw new Error("aborted");
    const useStream = forceStream || picked.workerMode === "stream";
    if (!useStream) {
      const initStart = performance.now();
      await loader.ensureLoaded();
      if (abortedError()) throw new Error("aborted");
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
        const { bytes, rowCount, droppedFeatures } = await exportAsStream(
          options.sheets,
          options.onProgress,
          options.signal,
        );
        result = {
          success: true,
          blob: new Blob([toBlobPart(bytes)], { type: XLSX_MIME }),
          // fast stream 是本包自带的纯 JS 写出器（fflate），不经 WASM 引擎
          engine: "fast-xlsx",
          mode: "stream",
          duration: performance.now() - start,
          rowCount,
          ...(droppedFeatures.length > 0 && {
            degradedFeatures: droppedFeatures,
          }),
        };
      } else {
        const builder = await WorkbookBuilder.create();
        if (abortedError()) throw new Error("aborted");
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
    // The Workbook build awaits toBuffer(), so an abort can arrive after the
    // pre-build check; do not report success/download for it.
    if (abortedError()) return aborted();
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
    if (abortedError()) return aborted();
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
      if (isAbort(e)) return aborted();
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
    if (abortedError()) return aborted();
    console.warn(
      `[excel-exporter] Worker path failed (${reason}); retrying on the main thread to preserve styles`,
    );
    try {
      return await runOnMainThread();
    } catch (e) {
      if (isAbort(e)) return aborted();
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
      if (isAbort(e)) return aborted();
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
      // A cancellation that raced the final worker response still wins: do
      // not report success/download for an export the caller aborted.
      if (abortedError()) return aborted();
      // Terminal 1 on success only: the failure route hands the sequence to
      // the degradation chain, which emits it exactly once (the types.ts
      // contract) -- emitting it here too duplicated the trailing 1.
      options.onProgress?.(1);
      if (options.download !== false) {
        triggerDownloadIsolated(options, result.blob!);
      }
      return result;
    }
    if (isAbort(result.error)) return aborted();
    // Worker export failed (e.g. WASM init error inside the worker) -> retry on
    // the main thread before degrading to the fast stream. Timeout-class
    // failures (noFallback, see worker-exporter) are the exception: the worker
    // was wedged or the workload legitimately exceeds workerTimeoutMs, so a
    // main-thread rerun of the same large build is exactly the freeze the
    // worker route exists to avoid (same policy as excel-preview's parse
    // timeout). Fail fast with the original error; the progress contract's
    // trailing 1 is emitted here since the degradation chain is skipped.
    if ((result.error as NoFallbackError | undefined)?.noFallback) {
      options.onProgress?.(1);
      return { ...result, duration: performance.now() - start };
    }
    return retryOnMainThread(result.error?.message ?? "worker export failed");
  } catch (e) {
    if (isAbort(e)) return aborted();
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
  // 转换期（tableExportToOptions）的结构化错误发生在进入 exportExcel 之前，
  // 不会落入其内部的结构化 catch——包一层让错误契约与主入口一致
  // （{success:false, error} 而非裸 reject），同一份坏输入走两条入口
  // 得到同一种失败形态。
  try {
    return await exportExcel(tableExportToOptions(options));
  } catch (error) {
    return { success: false, error: error as Error };
  }
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
  // 同 exportTable：转换期错误统一为结构化失败，不裸 reject。
  try {
    return await exportExcel(echartsExportToOptions(options));
  } catch (error) {
    return { success: false, error: error as Error };
  }
}
