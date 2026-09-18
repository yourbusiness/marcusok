import type { ColumnConfig, IndexColumnOptions, SheetConfig } from "./types";
import { columnProp, flattenColumnTree } from "./column-tree";

/**
 * prop of the virtual row-number column. Reserved: wherever a leaf column
 * carries this prop — injected by `applyIndexColumn` or written by hand on the
 * low-level builder paths — its cell values are generated from the row number
 * and never read from `data`. The reserved check goes through `columnProp`, so
 * the legacy `key` alias is reserved too (every other consumer reads the field
 * alias-aware; a prop-only check would treat `key: "__index__"` as a normal
 * column and silently read it from `data`).
 */
export const INDEX_PROP = "__index__";

/**
 * First number shown on the index column. Read from the (possibly boolean)
 * `indexColumn` config; the low-level builder paths without one start at 1.
 * Guards against `null` (`typeof null === "object"`) so a stray value on the
 * builder-direct paths degrades to start 1 instead of throwing mid-build.
 */
export function indexColumnStart(sheet: SheetConfig): number {
  const ic = sheet.indexColumn;
  return ic && typeof ic === "object" ? (ic.start ?? 1) : 1;
}

/**
 * Expand `sheet.indexColumn` into a real leading column: unshift a leaf column
 * with `INDEX_PROP` and shift every data-area merge one column right, so user
 * merges keep pointing at their original targets. Runs once at the
 * `exportExcel` entry (before mode routing), which makes the feature work on
 * every downstream path — worker and stream included — without per-builder
 * injection logic.
 */
export function applyIndexColumn(sheet: SheetConfig): SheetConfig {
  if (!sheet.indexColumn) return sheet;
  const opt: IndexColumnOptions =
    sheet.indexColumn === true ? {} : sheet.indexColumn;

  // 同名 prop 冲突宁可报错：序号列的值由行号生成、不读 data，用户列若撞名
  // 会被静默遮蔽成序号——数据悄悄丢失比报错更糟。
  const { leaves } = flattenColumnTree(sheet.columns);
  if (leaves.some((c) => columnProp(c) === INDEX_PROP)) {
    throw new Error(
      `[excel-exporter] column prop "${INDEX_PROP}" is reserved for the index column; rename your column or drop indexColumn`,
    );
  }

  const column: ColumnConfig = {
    prop: INDEX_PROP,
    label: opt.label ?? "序号",
    // 用 ?? 而非 ||：width 0 合法（隐藏列），不能被默认值 6 吞掉
    width: opt.width ?? 6,
    ...(opt.style !== undefined && { style: opt.style }),
    ...(opt.headerStyle !== undefined && { headerStyle: opt.headerStyle }),
  };

  return {
    ...sheet,
    columns: [column, ...sheet.columns],
    // merges 相对数据区 0 基定位，坐标基于用户自己的 columns；序号列插入后
    // 所有列右移一位，用户 merges 同步 +1 才仍指向原目标。
    merges: sheet.merges?.map((m) => ({ ...m, col: m.col + 1 })),
  };
}
