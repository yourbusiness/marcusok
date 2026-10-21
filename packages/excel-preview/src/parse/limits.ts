/**
 * 模型构建期的维度/几何校验。
 *
 * 背景见 src/limits.ts：文件声明可以远超 Excel 上限，直接进入布局会
 * 在主线程分配数 GB。这里把越界声明归一为 CORRUPT（与零 sheet 等畸形
 * 文件同一错误面），并把负/非有限的行高列宽归一为默认值。
 */
import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS } from "../limits";
import type { PreviewMerge } from "../types";

export function corruptSheet(message: string): never {
  throw Object.assign(new Error(message), { code: "CORRUPT" });
}

export function assertRowIndex(sheetName: string, index: number): void {
  if (!Number.isInteger(index) || index < 1 || index > EXCEL_MAX_ROWS) {
    corruptSheet(
      `Sheet "${sheetName}" declares row ${index}, outside Excel's 1..${EXCEL_MAX_ROWS} range. The file is corrupt or unsupported.`,
    );
  }
}

/** `col` is 0-based. */
export function assertColIndex(sheetName: string, col: number): void {
  if (!Number.isInteger(col) || col < 0 || col >= EXCEL_MAX_COLS) {
    corruptSheet(
      `Sheet "${sheetName}" declares column ${col + 1}, outside Excel's 1..${EXCEL_MAX_COLS} range. The file is corrupt or unsupported.`,
    );
  }
}

/** Column spans use 1-based inclusive min/max (the `<col min max>` form). */
export function assertColumnSpan(
  sheetName: string,
  min: number,
  max: number,
): void {
  if (
    !Number.isInteger(min) ||
    !Number.isInteger(max) ||
    min < 1 ||
    max < min ||
    max > EXCEL_MAX_COLS
  ) {
    corruptSheet(
      `Sheet "${sheetName}" declares column span ${min}..${max}, outside Excel's 1..${EXCEL_MAX_COLS} range. The file is corrupt or unsupported.`,
    );
  }
}

export function assertMerge(sheetName: string, merge: PreviewMerge): void {
  const endRow = merge.row + merge.rowSpan;
  const endCol = merge.col + merge.colSpan;
  if (
    !Number.isInteger(merge.row) ||
    !Number.isInteger(merge.col) ||
    !Number.isInteger(merge.rowSpan) ||
    !Number.isInteger(merge.colSpan) ||
    merge.row < 0 ||
    merge.col < 0 ||
    merge.rowSpan < 1 ||
    merge.colSpan < 1 ||
    endRow > EXCEL_MAX_ROWS ||
    endCol > EXCEL_MAX_COLS
  ) {
    corruptSheet(
      `Sheet "${sheetName}" declares a merged range outside Excel's ${EXCEL_MAX_ROWS} row / ${EXCEL_MAX_COLS} column limits. The file is corrupt or unsupported.`,
    );
  }
}

/** 非法（负/非有限）列宽回退默认值，避免几何变成负数/NaN。 */
export function normalizeWidth(width: number, fallback = 8.43): number {
  return Number.isFinite(width) && width >= 0 ? width : fallback;
}

/** 非法（负/非有限）行高回退默认行高（null = 15pt）。 */
export function normalizeRowHeight(
  height: number | null | undefined,
): number | null {
  return typeof height === "number" && Number.isFinite(height) && height >= 0
    ? height
    : null;
}

/** 冻结行列数钳制到 Excel 上限（负数/非整数视为 0）。 */
export function clampFrozen(
  value: number | null | undefined,
  max: number,
): number {
  return typeof value === "number" && Number.isInteger(value) && value > 0
    ? Math.min(value, max)
    : 0;
}
