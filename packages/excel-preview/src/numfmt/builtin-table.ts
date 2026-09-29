/**
 * 内置数字格式 id → Excel 实际渲染格式串（en-US 行为表）。
 *
 * 存在的原因（实测 2026-09）：ECMA-376 对内置 id 14-22 等定义的是标准串
 * （如 14 = "mm-dd-yy"），而 Excel 实际按 en-US 区域渲染为 "m/d/yyyy"
 * （MS-OE376 官方承认该实现差异）。modern-xlsx 的 getBuiltinFormat 返回
 * ECMA 标准串，且 loadFormatTable 覆盖在内置表恒优先的查表顺序下结构性
 * 无效（源码级核验：builtinFormatLookup 恒先于 CUSTOM_FORMATS）。唯一可行
 * 路径是格式化层自带"实际行为表"，把内置 id 先解析成格式串再交给引擎。
 *
 * 5-8 / 41-44 中的 _ 与 * 是宽度填充记号（引擎与补偿层都会丢弃），括号/
 * 减号字面量由补偿层重新注入——与 Excel 的会计格式视觉近似。
 */
export const EXCEL_BUILTIN_FORMATS: ReadonlyMap<number, string> = new Map([
  [0, "General"],
  [1, "0"],
  [2, "0.00"],
  [3, "#,##0"],
  [4, "#,##0.00"],
  [5, "$#,##0_);($#,##0)"],
  [6, "$#,##0_);[Red]($#,##0)"],
  [7, "$#,##0.00_);($#,##0.00)"],
  [8, "$#,##0.00_);[Red]($#,##0.00)"],
  [9, "0%"],
  [10, "0.00%"],
  [11, "0.00E+00"],
  [12, "# ?/?"],
  [13, "# ??/??"],
  // 14-22：locale 依赖的日期时间组，此处取 en-US 实际行为
  [14, "m/d/yyyy"],
  [15, "d-mmm-yy"],
  [16, "d-mmm"],
  [17, "mmm-yy"],
  [18, "h:mm AM/PM"],
  [19, "h:mm:ss AM/PM"],
  [20, "h:mm"],
  [21, "h:mm:ss"],
  [22, "m/d/yyyy h:mm"],
  // 23-36 为国际货币占位（现实中罕见）：不提供行为表，回退引擎的
  // getBuiltinFormat（ECMA 标准串），比没有强。
  [37, "#,##0_);(#,##0)"],
  [38, "#,##0_);[Red](#,##0)"],
  [39, "#,##0.00_);(#,##0.00)"],
  [40, "#,##0.00_);[Red](#,##0.00)"],
  [41, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'],
  [42, '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)'],
  [43, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'],
  [44, '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'],
  [45, "mm:ss"],
  [46, "[h]:mm:ss"],
  [47, "mm:ss.0"],
  [48, "##0.0E+0"],
  [49, "@"],
]);

/**
 * 解析 numFmtId → 格式串。
 * 自定义（>=164）优先查文件的 numFmts；内置查行为表；其余回退引擎标准串；
 * 一切失败兜底 "General"（formatCell 对 undefined/null 码会抛 TypeError，
 * 实测如此，样式链解析必须兜底）。
 */
export function resolveFormatCode(
  numFmtId: number,
  customNumFmts: ReadonlyMap<number, string>,
  engineBuiltin: (id: number) => string | undefined,
): string {
  if (numFmtId === 0) return "General";
  if (numFmtId >= 164) {
    const c = customNumFmts.get(numFmtId);
    if (c) return c;
  }
  const b = EXCEL_BUILTIN_FORMATS.get(numFmtId);
  if (b) return b;
  const e = engineBuiltin(numFmtId);
  if (e) return e;
  return "General";
}
