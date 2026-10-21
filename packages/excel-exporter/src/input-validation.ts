/**
 * exportExcel 的入口前置校验。
 *
 * 目标与历史注释一致：把"会产出 Excel 判损文件、或在两条路径上行为分裂"
 * 的输入挡在构建之前，以 `{ success: false, error }` 返回，而不是先降级
 * 再失败。样式枚举/XML 字符校验见 style-validation.ts / xml-safety.ts。
 */
import type { ColumnConfig, ExportMode, ExportOptions } from "./types";
import { columnLabel, flattenColumnTree } from "./column-tree";
import {
  validateMerges,
  validateSheetLimits,
  validateSheetName,
} from "./format-utils";
import { assertNoReservedIndexProp } from "./sheet-normalize";
import { validateCellStyle } from "./style-validation";
import { assertXmlSafeText } from "./xml-safety";

/** validateInput 的 mode 枚举校验用；与 ExportMode 类型一一对应。 */
const VALID_MODES = new Set<ExportMode>(["auto", "main", "worker", "stream"]);

/** 递归校验列树（含分组列）上的样式；分组列的 style 虽被忽略，非法值一并拦截。 */
function validateColumnStyles(
  columns: ColumnConfig[],
  sheetName: string,
): void {
  for (const col of columns) {
    const at = `sheet "${sheetName}" column "${columnLabel(col)}"`;
    validateCellStyle(col.style, `${at} style`);
    validateCellStyle(col.headerStyle, `${at} headerStyle`);
    if (col.children) validateColumnStyles(col.children, sheetName);
  }
}

/**
 * Pre-flight validation of user input. Runs the same checks as the
 * Workbook/stream build paths (same functions, same messages), hoisted
 * to the entry so a configuration error fails immediately with `{ success:
 * false, error }` instead of first degrading to a stream fallback attempt
 * that re-runs the identical checks and fails identically. Engine failures
 * (WASM unavailable, build errors) still degrade to the stream as before.
 */
export function validateInput(options: ExportOptions): void {
  // `options` 本身可能为 null/undefined（JS 调用方）：不守住这里，入口处在
  // try 之前解引用 options 的属性会以裸 TypeError reject，绕过下方承诺的
  // 结构化 { success: false } 契约。
  if (options === null || typeof options !== "object") {
    throw new Error("[excel-exporter] options must be an object");
  }
  // mode 是路由的显式意图输入：拼错的枚举值（"Main"/"STREAM"）会被 pickMode
  // 静默当作 auto 路由，用户的显式指定被无声忽略。与 filename/sheets 等
  // 字段同一口径前置报错；undefined = 未指定，走 auto。
  if (options.mode !== undefined && !VALID_MODES.has(options.mode)) {
    throw new Error(
      `[excel-exporter] mode must be one of "auto", "main", "worker", "stream" (got ${JSON.stringify(
        options.mode,
      )})`,
    );
  }
  // Guard the other core input alongside the sheets checks below: without it,
  // a JS caller omitting `filename` only fails inside triggerDownload with a
  // masked TypeError (caught as a cryptic warning), leaving success:true and
  // no file on disk. Fail fast with the structured { success: false } instead.
  if (typeof options.filename !== "string" || options.filename.length === 0) {
    throw new Error("[excel-exporter] filename must be a non-empty string");
  }
  // An empty sheets array is not a build error on the stream path (fast-xlsx
  // would zip a zero-sheet workbook Excel flags as corrupt while reporting
  // success), so reject it here like every other structural input error.
  // The same guard covers a missing/non-array `sheets`, which the caller's
  // totalRows reduce could otherwise hit first with a raw TypeError.
  if (!Array.isArray(options.sheets) || options.sheets.length === 0) {
    throw new Error("[excel-exporter] at least one sheet is required");
  }
  const seen = new Set<string>();
  for (const sheet of options.sheets) {
    // A null/primitive entry would fail on `sheet.name` with a raw TypeError;
    // give it the same clear treatment as the other structural guards below.
    if (sheet === null || typeof sheet !== "object") {
      throw new Error("[excel-exporter] each sheet must be an object");
    }
    validateSheetName(sheet.name);
    // Excel 的表名唯一性是大小写不敏感的（"Sheet1" 与 "SHEET1" 视为同名，
    // UI 层就拒绝这样命名），查重须同口径，否则会产出 Excel 可能要求修复的
    // 工作簿而 success 仍为 true。
    const nameKey = sheet.name.toLowerCase();
    if (seen.has(nameKey)) {
      throw new Error(`[excel-exporter] duplicate sheet name "${sheet.name}"`);
    }
    seen.add(nameKey);
    // Guard the two arrays the build paths index into; without these, a sheet
    // missing `columns`/`data` fails downstream with a raw TypeError instead
    // of a clear, actionable message.
    if (!Array.isArray(sheet.columns)) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" must have a columns array`,
      );
    }
    if (!Array.isArray(sheet.data)) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" must have a data array`,
      );
    }
    const { leaves, headerRowCount } = flattenColumnTree(sheet.columns);
    // 保留列在这里**无条件**拦截（与是否启用 indexColumn 无关）：构建器对
    // INDEX_PROP 的特判是无条件的，放行只会让用户自带的该列被行号静默顶替
    // （数据丢失而 success 仍为 true）。原先该检查只在 applyIndexColumn 内，
    // 未启用 indexColumn 时它直接返回原表，静默路径正好漏在最外层入口上。
    assertNoReservedIndexProp(leaves);
    validateMerges(sheet, leaves.length);
    // Excel 硬上限（1,048,576 行 × 16,384 列）：超限在 stream 路径会照常
    // 写出并 zip 成功、Excel 打开即报修复而 success 仍为 true。本校验在序号
    // 列注入之前运行，列数须自行加上 indexColumn 占用的一列。
    validateSheetLimits(
      sheet,
      leaves.length + (sheet.indexColumn ? 1 : 0),
      headerRowCount,
    );
    // 样式数值字段（textRotation / font.size）与枚举/XML 字符前置校验，
    // 覆盖 sheet 级、列级（含分组列）与 indexColumn 注入列的全部出现点。
    validateCellStyle(sheet.headerStyle, `sheet "${sheet.name}" headerStyle`);
    validateCellStyle(sheet.dataStyle, `sheet "${sheet.name}" dataStyle`);
    validateColumnStyles(sheet.columns, sheet.name);
    if (typeof sheet.indexColumn === "object" && sheet.indexColumn !== null) {
      validateCellStyle(
        sheet.indexColumn.style,
        `sheet "${sheet.name}" indexColumn style`,
      );
      validateCellStyle(
        sheet.indexColumn.headerStyle,
        `sheet "${sheet.name}" indexColumn headerStyle`,
      );
      if (
        sheet.indexColumn.label !== undefined &&
        typeof sheet.indexColumn.label !== "string"
      ) {
        throw new Error(
          `[excel-exporter] sheet "${sheet.name}" indexColumn.label must be a string`,
        );
      }
    }
    // 数值型布局/格式字段的前置校验：缺了这层，非法 width/freezeRows 会直通
    // 引擎并以晦涩的 serde 错误失败（"JSON parse error: invalid type: null,
    // expected f64"），随后整份导出被降级为无样式 stream；而同一输入在
    // >=50k 的 stream 路由却被静默忽略（width）——同一错误跨阈值一边降级
    // 一边成功。在此拦截使两条路径行为一致，且报错能定位到具体的列/表。
    if (
      sheet.freezeRows !== undefined &&
      (typeof sheet.freezeRows !== "number" ||
        !Number.isInteger(sheet.freezeRows) ||
        sheet.freezeRows < 0)
    ) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" freezeRows must be a non-negative integer`,
      );
    }
    // indexColumn 形状校验：非 boolean 非（非 null）object 的值（JS 调用方
    // 可能传 null / 字符串 / 数字）若放行，会被归一化与构建器各自按不同
    // 方式静默吞掉——与 freezeRows 等字段一致，前置报错。
    if (
      sheet.indexColumn !== undefined &&
      typeof sheet.indexColumn !== "boolean" &&
      (typeof sheet.indexColumn !== "object" || sheet.indexColumn === null)
    ) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" indexColumn must be a boolean or an options object`,
      );
    }
    // indexColumn.start 与 freezeRows 同类数值校验：非整数/负数只会在构建期
    // 产出 1.5、-1 这类怪序号，前置拦截。
    if (
      typeof sheet.indexColumn === "object" &&
      sheet.indexColumn !== null &&
      sheet.indexColumn.start !== undefined &&
      (!Number.isInteger(sheet.indexColumn.start) ||
        sheet.indexColumn.start < 0)
    ) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" indexColumn.start must be a non-negative integer`,
      );
    }
    // indexColumn.width 与下方 col.width 同类，但序号列是在**本校验之后**才由
    // applyIndexColumn 注入的（见 exportExcel），不在这里单独覆盖就会绕过校验：
    // NaN 直通引擎以 serde 错误失败并把整份导出降级为无样式 stream，而 >=50k 的
    // stream 路由又把 width 当作被丢弃的特性静默忽略——同一输入跨阈值一边降级
    // 一边成功，正是 col.width 校验要消除的那种分裂。
    if (
      typeof sheet.indexColumn === "object" &&
      sheet.indexColumn !== null &&
      sheet.indexColumn.width !== undefined &&
      (typeof sheet.indexColumn.width !== "number" ||
        !Number.isFinite(sheet.indexColumn.width) ||
        sheet.indexColumn.width < 0)
    ) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" indexColumn.width must be a finite non-negative number`,
      );
    }
    for (const col of leaves) {
      // OOXML 中 width=0 合法（隐藏列），负数与非有限数非法；只有叶子列
      // 消费 width，分组列上的 width 本就被忽略。
      if (
        col.width !== undefined &&
        (typeof col.width !== "number" ||
          !Number.isFinite(col.width) ||
          col.width < 0)
      ) {
        throw new Error(
          `[excel-exporter] column "${columnLabel(col)}" width must be a finite non-negative number`,
        );
      }
      if (col.format && typeof col.format === "object") {
        const spec = col.format;
        const at = `column "${columnLabel(col)}"`;
        // type 枚举校验：拼错的判别值（如 "percent"）会落进 applyFormat 的
        // default 分支被 toStr 字符串化——数字列整列变文本且 success:true
        // 零告警，正是 mode 枚举校验注释里要消灭的"显式指定被无声忽略"。
        // 先取到 string 变量再比较：直接对 spec.type 做全量不等比较会把
        // spec 收窄成 never，后续分支用不了。
        const specType: string = spec.type;
        if (
          specType !== "number" &&
          specType !== "date" &&
          specType !== "datetime" &&
          specType !== "enum" &&
          specType !== "padding"
        ) {
          throw new Error(
            `[excel-exporter] ${at} format.type must be one of "number", "date", "datetime", "enum", "padding" (got ${JSON.stringify(specType)})`,
          );
        }
        // enum.map 为必填对象：缺失/null/数组会在渲染期以晦涩 TypeError
        // （Object.hasOwn 目标非法）直通降级链——与 decimals/length 同类，
        // 前置拦截且报错能定位到列。
        if (
          spec.type === "enum" &&
          (typeof spec.map !== "object" ||
            spec.map === null ||
            Array.isArray(spec.map))
        ) {
          throw new Error(
            `[excel-exporter] ${at} format.map must be an object mapping values to strings`,
          );
        }
        // date/datetime.pattern：非字符串在 formatDateByPattern 的
        // toLowerCase 即抛裸 TypeError（stream 路径 displayValue 消费）；
        // XML 非法字符会被原样写进 numFmt 的 formatCode，产出 Excel 判损
        // 的 styles.xml（Workbook 路径），因此与 font.name/numFormat 同口径
        // 前置拒绝。
        if (
          (spec.type === "date" || spec.type === "datetime") &&
          spec.pattern !== undefined
        ) {
          assertXmlSafeText(spec.pattern, at, "format.pattern");
        }
        // padding.fill：非字符串不抛错但产出怪输出（padStart 把 fill 先
        // String() 化再取首字符），与 length 校验同类前置；类型上必填，
        // undefined 一并拦截。控制字符属于数据文本，构建期会统一清洗。
        if (spec.type === "padding" && typeof spec.fill !== "string") {
          throw new Error(
            `[excel-exporter] ${at} format.fill must be a string`,
          );
        }
        // decimals 两路径共享：Workbook 路径拼 numFormat 字符串（任意值都
        // "能出"），stream 路径烧入 toFixed(decimals)——收敛到 toFixed 自身
        // 的 0..100 上限，同一 spec 才不会在 50k 行上下一边成功一边抛错。
        if (
          spec.type === "number" &&
          spec.decimals !== undefined &&
          (!Number.isInteger(spec.decimals) ||
            spec.decimals < 0 ||
            spec.decimals > 100)
        ) {
          throw new Error(
            `[excel-exporter] ${at} format.decimals must be an integer between 0 and 100`,
          );
        }
        // padding.length：非整数/负数会让 padStart 抛 RangeError 或静默不
        // 填充，超大值会生成巨型字符串，均在渲染期才爆——前置拦截。
        if (
          spec.type === "padding" &&
          (!Number.isInteger(spec.length) ||
            spec.length < 0 ||
            spec.length > 10_000)
        ) {
          throw new Error(
            `[excel-exporter] ${at} format.length must be an integer between 0 and 10000`,
          );
        }
      }
    }
  }
}
