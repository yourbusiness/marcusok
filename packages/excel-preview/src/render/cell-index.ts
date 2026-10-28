/**
 * 行内单元格索引：按列有序的扁平数组 + 二分查找。
 *
 * 背景（性能）：窗口差量渲染此前用 per-row Map 全量遍历再按列过滤——宽表
 * （行内上万非空格）下每个滚动帧都要扫"窗口行数 × 行内全部格"，横向滚动
 * 的帧预算被窗口外的列吃掉。模型（PreviewRow.cells）不可变且通常已按列
 * 升序，setSheet 时一次性构建有序数组，窗口循环与点查都走二分：每帧成本
 * 从 O(行数 × 行内总格数) 降到 O(行数 × (log 行内格数 + 窗口格数))。
 *
 * 语义与旧 Map 形态对齐：
 * - 同列重复（引擎正常产物不含，防御性处理）保留后一个——与 Map.set 的
 *   覆盖语义一致；
 * - 迭代按列升序。旧形态为引擎文档序，两者只在异常文件的重叠格上有 DOM
 *   追加顺序差异——格子互不重叠，无视觉影响。
 */
import type { PreviewCell, PreviewRow } from "../types";

/** 行号(0-based) → 该行按列升序的单元格数组。 */
export type CellRowIndex = Map<number, PreviewCell[]>;

/** 从模型行构建索引：复制 cells，乱序才排序、同列去重保末位。 */
export function buildRowIndex(rows: PreviewRow[]): CellRowIndex {
  const index: CellRowIndex = new Map();
  for (const row of rows) {
    const cells = row.cells.slice();
    // 升序校验（O(n)，正常文件即引擎文档序已有序，零排序成本）
    let ascending = true;
    for (let i = 1; i < cells.length; i++) {
      if (cells[i].col < cells[i - 1].col) {
        ascending = false;
        break;
      }
    }
    // sort 稳定（V8 保证）：同列保持原有相对顺序，去重保末位即 Map.set 语义
    if (!ascending) cells.sort((a, b) => a.col - b.col);
    // 同列重复的判定必须在排序之后做：乱序输入里的重复对（如 5,2,5,0）
    // 排序前不相邻，相邻比较会漏判
    let dup = false;
    for (let i = 1; i < cells.length; i++) {
      if (cells[i].col === cells[i - 1].col) {
        dup = true;
        break;
      }
    }
    if (dup) {
      const dedup: PreviewCell[] = [];
      for (const c of cells) {
        if (dedup.length && dedup[dedup.length - 1].col === c.col) {
          dedup[dedup.length - 1] = c;
        } else {
          dedup.push(c);
        }
      }
      index.set(row.index - 1, dedup);
      continue;
    }
    index.set(row.index - 1, cells);
  }
  return index;
}

/** 二分：cells（按列升序）中首个 col >= target 的下标。 */
function lowerBound(cells: PreviewCell[], target: number): number {
  let lo = 0;
  let hi = cells.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (cells[mid].col < target) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** 点查（旧 Map.get 的等价物）：命中返回格子，缺失返回 undefined。 */
export function cellAt(
  cells: PreviewCell[] | undefined,
  col: number,
): PreviewCell | undefined {
  if (!cells) return undefined;
  const i = lowerBound(cells, col);
  return i < cells.length && cells[i].col === col ? cells[i] : undefined;
}

/** 窗口迭代：按列升序回调 [colStart, colEnd) 内的全部格子。 */
export function forEachCellInRange(
  cells: PreviewCell[] | undefined,
  colStart: number,
  colEnd: number,
  cb: (cell: PreviewCell) => void,
): void {
  if (!cells || colStart >= colEnd) return;
  for (let i = lowerBound(cells, colStart); i < cells.length; i++) {
    const cell = cells[i];
    if (cell.col >= colEnd) return;
    cb(cell);
  }
}
