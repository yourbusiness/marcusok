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
  /** 全部合并（主格召回判定用，见 visibleMergeAnchors）。 */
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
    // 隐藏列宽 0；<col> 出现即采用其 width——引擎对未写 width 的 col 兜底
    // 8.43（与默认一致，无副作用），此前设 customWidth 门槛会把"只写 width
    // 不写标志"的第三方产物列宽整体吞掉（实测引擎两形态都给 customWidth
    // =false，无从区分也不必区分；customWidth 字段在模型里保留仅作元信息）
    const w = charsToPx(span.width);
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

/** 可视范围 + 缓冲（纯窗口，不含合并召回——见 visibleMergeAnchors）。 */
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

  const r0 = Math.min(rowStart, layout.rowCount - 1);
  const r1 = Math.min(Math.max(rowEnd, r0 + 1), layout.rowCount);
  const c0 = Math.min(colStart, layout.colCount - 1);
  const c1 = Math.min(Math.max(colEnd, c0 + 1), layout.colCount);

  return { rowStart: r0, rowEnd: r1, colStart: c0, colEnd: c1 };
}

/**
 * 与主区窗口相交的合并的主格（anchor）列表：渲染层据此**增补渲染**窗口外的
 * 合并主格（主格的 width/height 直接按合并跨度跨越，覆盖格不逐格建 DOM，
 * 视口内的合并因此显示完整）。
 *
 * 此前召回是把整个合并矩形并入窗口 [r0,r1)/[c0,c1)——窗口与行号/列标表头
 * 都按扩大后的范围循环建 DOM，A1:A100000 一类整列合并（Excel 模板常态）
 * 会把行号表头全量建出（10 万~百万 DOM），虚拟滚动失效。改为只召回主格后
 * 表头恒为视口规模。冻结区象限不需要召回：其行/列范围本就是 [0, frozen)
 * 全量渲染，anchor 落在里面自然被覆盖（跨冻结线的异常合并除外，边角行为
 * 与旧实现等同）。
 */
export function visibleMergeAnchors(
  layout: GridLayout,
  range: VisibleRange,
): { row: number; col: number }[] {
  const out: { row: number; col: number }[] = [];
  for (const m of layout.merges) {
    if (m.row < layout.frozenRows || m.col < layout.frozenCols) continue;
    const endR = m.row + m.rowSpan;
    const endC = m.col + m.colSpan;
    if (endR <= range.rowStart || m.row >= range.rowEnd) continue;
    if (endC <= range.colStart || m.col >= range.colEnd) continue;
    out.push({ row: m.row, col: m.col });
  }
  return out;
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
