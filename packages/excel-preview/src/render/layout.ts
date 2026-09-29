/**
 * 网格布局几何：列宽/行高归一、前缀和偏移、可视窗口计算（虚拟滚动的数学层）。
 *
 * 单位换算（Excel 默认 Calibri 11 @ 96dpi）：
 *  - 列宽字符单位 → 像素：round(width × 7 + 5)（MDW=7 的标准近似）
 *  - 行高 pt → 像素：round(pt × 4 / 3)
 */
import type { PreviewMerge, PreviewSheet } from "../types";

export const DEFAULT_COL_WIDTH_CHARS = 8.43;
export const DEFAULT_ROW_HEIGHT_PT = 15;
const MDW = 7; // max digit width, Calibri 11 @ 96dpi
const HEAD_ROW_PX = 24; // 表头（列标）行高
const HEAD_COL_PX = 40; // 表头（行号）列宽

export function charsToPx(width: number): number {
  return Math.round(width * MDW + 5);
}

export function ptToPx(pt: number): number {
  return Math.round((pt * 4) / 3);
}

/** 网格几何缓存：由 renderer 在 sheet/尺寸变化时重建。 */
export interface GridLayout {
  /** 每列像素宽（0-based，下标即列号）。 */
  colWidths: Float64Array;
  /** 每行像素高（0-based，下标即行号）。 */
  rowHeights: Float64Array;
  /** 前缀和：colLeft[i] = 第 i 列左缘；长度 = 列数 + 1。 */
  colLeft: Float64Array;
  rowTop: Float64Array;
  totalWidth: number;
  totalHeight: number;
  /** 合并索引：anchor `row:col` → merge。 */
  mergeByAnchor: Map<string, PreviewMerge>;
  /** 被（非自身）合并覆盖的格子：`row:col` → anchor key。 */
  coveredBy: Map<string, string>;
  /** 全部合并（可视窗口合并召回用）。 */
  merges: PreviewMerge[];
  /** 内容边界（0-based，不含表头偏移）。 */
  rowCount: number;
  colCount: number;
  frozenRows: number;
  frozenCols: number;
}

function buildPrefix(values: Float64Array): Float64Array {
  const pre = new Float64Array(values.length + 1);
  for (let i = 0; i < values.length; i++) pre[i + 1] = pre[i] + values[i];
  return pre;
}

export function buildLayout(sheet: PreviewSheet): GridLayout {
  const colCount = Math.max(sheet.colCount, sheet.frozenCols, 1);
  const rowCount = Math.max(sheet.rowCount, sheet.frozenRows, 1);

  const colWidths = new Float64Array(colCount);
  colWidths.fill(charsToPx(DEFAULT_COL_WIDTH_CHARS));
  for (const span of sheet.colSpans) {
    // 隐藏列宽 0；customWidth 才覆盖默认
    const w = span.hidden
      ? 0
      : span.customWidth
        ? charsToPx(span.width)
        : charsToPx(DEFAULT_COL_WIDTH_CHARS);
    for (let c = span.min - 1; c <= span.max - 1 && c < colCount; c++) {
      colWidths[c] = span.hidden ? 0 : w;
    }
  }

  const rowHeights = new Float64Array(rowCount);
  rowHeights.fill(ptToPx(DEFAULT_ROW_HEIGHT_PT));
  for (const r of sheet.rows) {
    const i = r.index - 1;
    if (i >= 0 && i < rowCount) {
      rowHeights[i] = r.hidden ? 0 : ptToPx(r.height ?? DEFAULT_ROW_HEIGHT_PT);
    }
  }

  const colLeft = buildPrefix(colWidths);
  const rowTop = buildPrefix(rowHeights);

  const mergeByAnchor = new Map<string, PreviewMerge>();
  const coveredBy = new Map<string, string>();
  for (const m of sheet.merges) {
    const anchor = `${m.row}:${m.col}`;
    mergeByAnchor.set(anchor, m);
    for (let r = m.row; r < m.row + m.rowSpan; r++) {
      for (let c = m.col; c < m.col + m.colSpan; c++) {
        if (r === m.row && c === m.col) continue;
        coveredBy.set(`${r}:${c}`, anchor);
      }
    }
  }

  return {
    colWidths,
    rowHeights,
    colLeft,
    rowTop,
    totalWidth: colLeft[colCount],
    totalHeight: rowTop[rowCount],
    mergeByAnchor,
    coveredBy,
    merges: sheet.merges,
    rowCount,
    colCount,
    frozenRows: sheet.frozenRows,
    frozenCols: sheet.frozenCols,
  };
}

/** 二分：给定偏移找第一个 >= v 的下标（返回 0-based 区间下标）。 */
function lowerBound(pre: Float64Array, v: number): number {
  let lo = 0;
  let hi = pre.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (pre[mid] < v) lo = mid + 1;
    else hi = mid;
  }
  return Math.max(0, lo - 1);
}

export interface VisibleRange {
  /** 含缓冲区的 [start, end) 半开区间（0-based 行列）。 */
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

/** 可视范围 + 缓冲；再并入与窗口相交的合并区（合并召回）。 */
export function computeVisibleRange(
  layout: GridLayout,
  scrollTop: number,
  scrollLeft: number,
  viewW: number,
  viewH: number,
  bufferPx = 200,
): VisibleRange {
  const rowStart = lowerBound(layout.rowTop, Math.max(0, scrollTop - bufferPx));
  const rowEnd = lowerBound(layout.rowTop, scrollTop + viewH + bufferPx) + 1;
  const colStart = lowerBound(
    layout.colLeft,
    Math.max(0, scrollLeft - bufferPx),
  );
  const colEnd = lowerBound(layout.colLeft, scrollLeft + viewW + bufferPx) + 1;

  let r0 = Math.min(rowStart, layout.rowCount - 1);
  let r1 = Math.min(Math.max(rowEnd, r0 + 1), layout.rowCount);
  let c0 = Math.min(colStart, layout.colCount - 1);
  let c1 = Math.min(Math.max(colEnd, c0 + 1), layout.colCount);

  // 合并召回：主单元格在窗口上方/左侧的跨行/跨列合并，其覆盖区只要与窗口
  // 相交就必须整体渲染（否则视口内的合并显示残缺）。合并表通常很小，线性
  // 扫描即可；超大合并文件下限为窗口本身（召回一次到位，无递归）。
  for (const m of layout.merges) {
    const endR = m.row + m.rowSpan;
    const endC = m.col + m.colSpan;
    if (endR <= r0 || m.row >= r1) continue;
    if (endC <= c0 || m.col >= c1) continue;
    if (m.row < r0) r0 = m.row;
    if (endR > r1) r1 = endR;
    if (m.col < c0) c0 = m.col;
    if (endC > c1) c1 = endC;
  }

  return { rowStart: r0, rowEnd: r1, colStart: c0, colEnd: c1 };
}

export const HEADER_ROW_PX = HEAD_ROW_PX;
export const HEADER_COL_PX = HEAD_COL_PX;

/** 0-based 列号 → 字母标头（A, B, …, Z, AA, AB…）。 */
export function columnIndexLabel(col: number): string {
  let s = "";
  let c = col + 1;
  while (c > 0) {
    const rem = (c - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}
