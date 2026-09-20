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
 *
 * `exportExcel` 对**用户自带的**该 prop 一律报错，与是否启用 `indexColumn`
 * 无关（见 assertNoReservedIndexProp）。低层构建器仍保留"手写即行号列"的
 * 语义——它们不做输入校验，调用方本就该知道自己在写保留列。
 */
export const INDEX_PROP = "__index__";

/**
 * 保留列冲突检查。构建器（`WorkbookBuilder.addSheet` / `fast-xlsx`）对
 * INDEX_PROP 的特判是**无条件**的：只要叶子列的 prop（或 legacy `key`）命中，
 * 该列的值就由行号生成、永不读 `data`。因此用户自带的 `__index__` 列不是
 * "报错"而是"数据被静默顶替、success 仍为 true"。
 *
 * `exportExcel`（validateInput）与 `applyIndexColumn` 共用本判定——原先该检查
 * 只存在于 `applyIndexColumn` 内，而它在未启用 `indexColumn` 时直接返回原表，
 * 于是那条静默路径正好漏在最外层入口上。主入口下用户不可能"有意"手写该列
 * （要序号列应声明 `indexColumn`），故报错不会误伤。
 */
export function assertNoReservedIndexProp(leaves: ColumnConfig[]): void {
  if (leaves.some((c) => columnProp(c) === INDEX_PROP)) {
    throw new Error(
      `[excel-exporter] column prop "${INDEX_PROP}" is reserved for the index column; rename your column or drop indexColumn`,
    );
  }
}

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
  // 会被静默遮蔽成序号——数据悄悄丢失比报错更糟。（exportExcel 路径上
  // validateInput 已先拦一次；此处保证直接调用本函数的调用方同样拿到报错。）
  assertNoReservedIndexProp(flattenColumnTree(sheet.columns).leaves);

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
