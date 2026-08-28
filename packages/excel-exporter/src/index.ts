import type { ExportOptions, ExportResult, ExportMode } from "./types";
import { WorkbookBuilder } from "./workbook-builder";
import { exportAsStream } from "./streaming-builder";
import { exportInWorker } from "./worker-exporter";
import { exportWithSheetJS } from "./fallback";
import { triggerDownload, toBlobPart } from "./download";
import { getWasmLoader } from "./wasm-loader";
import { tableExportToOptions, type TableExportOptions } from "./table-export";
import {
  echartsExportToOptions,
  type EChartsExportOptions,
} from "./echarts-export";
import { validateSheetName, validateMerges } from "./format-utils";
import { flattenColumnTree } from "./column-tree";

export * from "./types";
export * from "./style-presets";
export * from "./format-utils";
export * from "./table-export";
export * from "./echarts-export";
export { configureWasm, getWasmLoader } from "./wasm-loader";
export type { LoaderOptions, LoadState } from "./wasm-loader";
export { WorkbookBuilder } from "./workbook-builder";
export { exportAsStream } from "./streaming-builder";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const STREAM_THRESHOLD = 50_000; // Workbook.toBuffer cliff starts ~55k rows
const WORKER_THRESHOLD = 20_000; // main-mode sync work is acceptable below this

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
    // instead of silently degrading to the style-less SheetJS fallback.
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
 * Pre-flight validation of user input. Runs the same checks as the
 * Workbook/stream/SheetJS build paths (same functions, same messages), hoisted
 * to the entry so a configuration error fails immediately with `{ success:
 * false, error }` instead of first degrading to a SheetJS fallback attempt
 * that re-runs the identical checks and fails identically. Engine failures
 * (WASM unavailable, build errors) still degrade to SheetJS as before.
 */
function validateInput(options: ExportOptions): void {
  const seen = new Set<string>();
  for (const sheet of options.sheets) {
    validateSheetName(sheet.name);
    if (seen.has(sheet.name)) {
      throw new Error(`[excel-exporter] duplicate sheet name "${sheet.name}"`);
    }
    seen.add(sheet.name);
    const { leaves } = flattenColumnTree(sheet.columns);
    validateMerges(sheet, leaves.length);
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
 *       { key: 'product', header: 'Product', width: 20 },
 *       { key: 'revenue', header: 'Revenue', width: 15, style: StylePresets.currency },
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
  const totalRows = options.sheets.reduce((s, sh) => s + sh.data.length, 0);

  // Leading 0 fires exactly once here, on every route (the SheetJS fallback
  // included), so consumers always see the documented 0 -> ... -> 1 pair.
  options.onProgress?.(0);

  // Invalid input fails here on every route (same messages as before; the
  // build paths keep their own checks for direct callers). The trailing 1 is
  // still emitted so the 0 -> 1 progress contract holds for failed exports.
  try {
    validateInput(options);
  } catch (e) {
    options.onProgress?.(1);
    return {
      success: false,
      error: e as Error,
      duration: performance.now() - start,
    };
  }

  // The SheetJS fallback never reports progress itself; closing the sequence
  // here keeps the terminal-1 contract true on degraded routes too, including
  // when the fallback itself fails and resolves with success: false.
  const finishWithSheetJS = (reason: string): Promise<ExportResult> =>
    exportWithSheetJS(options, start, reason).finally(() =>
      options.onProgress?.(1),
    );

  const picked = pickMode(options, totalRows);
  const needsWasm = picked.workerMode !== "stream";
  const loader = getWasmLoader();
  if (needsWasm && !loader.supported) {
    return finishWithSheetJS("WebAssembly not supported");
  }

  // Node main/stream: execute directly on this thread (no Web Worker available).
  if (
    picked.mode === "main" ||
    (picked.mode === "stream" && typeof window === "undefined")
  ) {
    try {
      if (needsWasm) {
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
        if (picked.workerMode === "stream") {
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
        // back to SheetJS still shows how long it spent before failing.
        options.onPhase?.("build", performance.now() - buildStart);
      }
      options.onProgress?.(1);
      // Node has no document: triggerDownload would be a no-op, so neither the
      // click nor the "download" phase is reported (matches ExportPhase docs).
      if (options.download !== false && typeof document !== "undefined") {
        const downloadStart = performance.now();
        triggerDownload(result.blob!, options.filename);
        options.onPhase?.("download", performance.now() - downloadStart);
      }
      return result;
    } catch (e) {
      return finishWithSheetJS((e as Error).message);
    }
  }

  // Browser worker mode: offload to worker (main thread does one structured clone).
  try {
    const result = await exportInWorker(options, picked.workerMode!);
    if (result.success) {
      // Terminal 1 on success only: the failure route hands the sequence to
      // finishWithSheetJS, whose finally emits it exactly once (the types.ts
      // contract) -- emitting it here too duplicated the trailing 1.
      options.onProgress?.(1);
      if (options.download !== false) {
        const downloadStart = performance.now();
        triggerDownload(result.blob!, options.filename);
        options.onPhase?.("download", performance.now() - downloadStart);
      }
      return result;
    }
    // Worker export failed (e.g. WASM init error inside the worker) -> degrade
    // to SheetJS, matching the main-thread path's failure handling.
    return finishWithSheetJS(result.error?.message ?? "worker export failed");
  } catch (e) {
    return finishWithSheetJS((e as Error).message);
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
