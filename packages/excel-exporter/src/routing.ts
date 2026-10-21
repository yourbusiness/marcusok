/**
 * 自动路由（auto / 环境回退）。
 *
 * 历史阈值只看行数，但 Workbook 路径的成本由**单元格总量**决定：
 * 40k 行 × 50 列（200 万格）实测 6.3s、RSS +1.2GB，而 60k 行 × 20 列
 * （120 万格）走 stream 只需 978ms / +58MB。行数相同、列数不同的两张表
 * 会被路由到相差一个数量级的路径。因此阈值同时看行数与估算单元格数：
 * - stream 阈值：行数或单元格数任一达到即走 Fast stream（牺牲样式换吞吐）；
 * - worker 阈值：浏览器内行数或单元格数任一达到即交给 Worker，避免主线程
 *   同步构建阻塞 UI。
 *
 * 单元格预算的取值依据（本地实测，Node 24.21）：
 * - Workbook 2M 格：20k×100 3.2s/258MB、40k×50 6.3s/1.2GB；
 * - stream 2M 格：约 1s/100MB；
 * - 1.5M 格给 20k×50（1M 格，Workbook 约 1.6s）留出余量，同时把 2M 格
 *   的宽表挡在 Workbook 之外。
 */
import type { ExportMode, ExportOptions } from "./types";
import { flattenColumnTree } from "./column-tree";

export const STREAM_ROW_THRESHOLD = 50_000;
export const STREAM_CELL_THRESHOLD = 1_500_000;
export const WORKER_ROW_THRESHOLD = 20_000;
export const WORKER_CELL_THRESHOLD = 250_000;

export interface PickedMode {
  mode: ExportMode;
  workerMode?: "workbook" | "stream";
}

export function isBrowserEnv(): boolean {
  return typeof Worker !== "undefined" && typeof window !== "undefined";
}

/**
 * 估算总单元格数（含 indexColumn：调用点必须在 applyIndexColumn 之后）。
 * 稀疏数据也按"行 × 叶子列"计——构建成本由遍历规模决定，不取决于空值。
 */
export function estimateTotalCells(sheets: ExportOptions["sheets"]): number {
  let total = 0;
  for (const sheet of sheets) {
    total += sheet.data.length * flattenColumnTree(sheet.columns).leaves.length;
  }
  return total;
}

export function pickMode(
  mode: ExportMode | undefined,
  totalRows: number,
  totalCells: number,
  browser = isBrowserEnv(),
): PickedMode {
  const stream =
    totalRows >= STREAM_ROW_THRESHOLD || totalCells >= STREAM_CELL_THRESHOLD;
  const needsWorker =
    totalRows >= WORKER_ROW_THRESHOLD || totalCells >= WORKER_CELL_THRESHOLD;
  const explicit = mode ?? "auto";

  if (explicit === "stream") return { mode: "stream", workerMode: "stream" };
  if (explicit === "main") return { mode: "main" };
  if (explicit === "worker") {
    // Worker mode requires a Web Worker global. In environments without one
    // (Node/SSR), fall back to the main-thread path so styles are preserved
    // instead of silently degrading to the style-less stream fallback.
    if (!browser) {
      return stream
        ? { mode: "stream", workerMode: "stream" }
        : { mode: "main" };
    }
    return {
      mode: "worker",
      workerMode: stream ? "stream" : "workbook",
    };
  }

  // auto
  if (!browser) {
    return stream ? { mode: "stream", workerMode: "stream" } : { mode: "main" };
  }
  if (!needsWorker) return { mode: "main" };
  return {
    mode: "worker",
    workerMode: stream ? "stream" : "workbook",
  };
}
