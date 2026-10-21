/**
 * Excel/OOXML 的硬上限（ECMA-376）：每表 1,048,576 行（含表头/空行）与
 * 16,384 列（XFD）。
 *
 * 预览此前直接信任文件声明的 rowCount/colCount：一个几 KB 的畸形文件
 * （单元格 A100000000、合并 A1:A100000000、<col max="100000000">）会让
 * buildLayout 在**主线程**按声明维度分配数个 GB 的 TypedArray——worker
 * 超时保护覆盖不到布局阶段。模型构建阶段按本上限拒绝，布局层再做一次
 * 防御性钳制（低层调用方可能直接构造 PreviewSheet）。
 */
export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLS = 16_384;
