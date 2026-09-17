/**
 * 样式演示的预览渲染纯函数：把 CellStyle 映射为 CSS、把 numFormat 映射为
 * 浏览器内可见文本，并按导出引擎同样的规则解析"生效样式"。
 *
 * 预览保真原则：生效样式的解析顺序镜像 workbook-builder.applyLayout——
 * 表头是列级 headerStyle 整体替换表级默认，数据格是 dataStyle 与列级
 * style 字段级深合并——其中深合并直接引用库内真实现（见下方 mergeStyles
 * 导入注释），保证预览与真实导出永远共用同一套合并逻辑。
 */
import type { CSSProperties } from "react";
import type {
  BorderStyle,
  CellStyle,
  ColumnConfig,
  SheetConfig,
} from "@marcusok/excel-exporter";
// 跨包相对导入库内源码的深合并真实现：mergeStyles 未从公共入口导出，而 play
// 的 vitest 配置没有 @marcusok 别名（走公共入口会把 dist 全量引擎拉进 node
// 测试）。style-utils.ts 运行时零依赖，相对导入在 vite / vitest / tsc
// （bundler 解析 .js -> .ts）三端均可解析。将来库若公开导出 mergeStyles，
// 应改回公共入口导入。
import {
  mergeStyles,
  BaseCellStyle,
} from "../../../../excel-exporter/src/style-utils.js";

/** 复刻库内 sheet-normalize.INDEX_PROP（未从公共入口导出）。 */
const PREVIEW_INDEX_PROP = "__index__";

// ---------------------------------------------------------------------------
// 生效样式解析（镜像 workbook-builder.applyLayout 的应用顺序）
// ---------------------------------------------------------------------------

/**
 * 展开表级 indexColumn 为最左虚拟列，镜像库内 sheet-normalize.applyIndexColumn
 * 的可见行为（label/width 默认值、样式透传）。演示场景无 merges，省略 merges
 * 右移；也不做 __index__ 保留字冲突检查（demo 数据可控）。
 */
export function expandIndexColumn(sheet: SheetConfig): ColumnConfig[] {
  if (!sheet.indexColumn) return sheet.columns;
  const opt = sheet.indexColumn === true ? {} : sheet.indexColumn;
  const column: ColumnConfig = {
    prop: PREVIEW_INDEX_PROP,
    label: opt.label ?? "序号",
    // 用 ?? 而非 ||：width 0 合法（隐藏列），不能被默认值 6 吞掉——与库一致
    width: opt.width ?? 6,
    ...(opt.style !== undefined && { style: opt.style }),
    ...(opt.headerStyle !== undefined && { headerStyle: opt.headerStyle }),
  };
  return [column, ...sheet.columns];
}

/** 序号列首行数字，镜像库内 indexColumnStart。 */
function indexColumnStart(sheet: SheetConfig): number {
  const ic = sheet.indexColumn;
  return ic && typeof ic === "object" ? (ic.start ?? 1) : 1;
}

/**
 * FormatSpec 自动注入的 numFormat：镜像 format-utils.numFormatForSpec 的
 * date/datetime/number 三分支（enum/padding 产出纯字符串，无需注入）。
 * 不能直接导入库内实现——format-utils 有 modern-xlsx 运行时依赖，会把引擎
 * 拉进 node 测试，故在此镜像并用测试锁定行为。
 */
function numFormatForSpecMirror(col: ColumnConfig): string | null {
  const spec = typeof col.format === "object" ? col.format : null;
  if (!spec) return null;
  switch (spec.type) {
    case "date":
      return spec.pattern ?? "yyyy-MM-dd";
    case "datetime":
      return spec.pattern ?? "yyyy-MM-dd HH:mm";
    case "number": {
      const dec = spec.decimals ?? 0;
      const head = spec.thousands ? "#,##0" : "0";
      return dec > 0 ? `${head}.${"0".repeat(dec)}` : head;
    }
    default:
      return null;
  }
}

/**
 * 表头格生效样式：列级 headerStyle 整体替换表级默认（不合并），
 * 镜像 workbook-builder.ts 的 `cell.column.headerStyle ?? config.headerStyle`；
 * 基底 BaseCellStyle 铺在结果之下，未声明对齐的表头因此同样居中。
 */
export function effectiveHeaderStyle(
  sheet: SheetConfig,
  col: ColumnConfig,
): CellStyle | undefined {
  return mergeStyles(BaseCellStyle, col.headerStyle ?? sheet.headerStyle);
}

/**
 * 数据格生效样式：mergeStyles(dataStyle, 列样式)。注入检查点与库内
 * withAutoNumFormat 一致——只看列自身 style.numFormat 是否缺省，缺省时才
 * 注入 FormatSpec 对应的 numFormat；注入后的样式作为 override 参与深合并，
 * 因此列级 numFormat 会赢过 dataStyle 的（与导出一致）。
 */
export function effectiveDataStyle(
  sheet: SheetConfig,
  col: ColumnConfig,
): CellStyle | undefined {
  let override = col.style;
  if (override?.numFormat === undefined) {
    const nf = numFormatForSpecMirror(col);
    if (nf !== null) override = { ...(override ?? {}), numFormat: nf };
  }
  return mergeStyles(BaseCellStyle, mergeStyles(sheet.dataStyle, override));
}

/** 预览列模型：一列 = 表头样式 + 数据样式 + 取值所需的元信息。 */
export interface PreviewColumn {
  /** 数据字段名（index 列为 __index__）。同一 prop 可配多列，故不能直接当 React key。 */
  prop: string;
  /** 渲染用唯一 key：prop + 列序号（spread 派生场景三列同为 "amount"）。 */
  id: string;
  label: string;
  widthPx: number | undefined; // Excel 字符宽 × 7.5 的 px 近似
  headerStyle: CellStyle | undefined;
  dataStyle: CellStyle | undefined;
  isIndex: boolean;
}

/**
 * 把 SheetConfig 解析为预览列。演示场景均为扁平表头（无 children），顶层列
 * 即叶子列，无需展开列树。
 */
export function buildPreviewColumns(sheet: SheetConfig): PreviewColumn[] {
  return expandIndexColumn(sheet).map((col, index) => {
    const prop = col.prop ?? "";
    return {
      prop,
      id: `${prop}#${index}`,
      label: col.label ?? col.prop ?? "",
      widthPx:
        col.width !== undefined ? Math.round(col.width * 7.5) : undefined,
      headerStyle: effectiveHeaderStyle(sheet, col),
      dataStyle: effectiveDataStyle(sheet, col),
      isIndex: prop === PREVIEW_INDEX_PROP,
    };
  });
}

/**
 * 预览单元格显示文本：序号列由行号生成（不读 data，与库一致）；有
 * numFormat 时按格式码渲染（数字对数字码、ISO 字符串对日期码）；其余
 * String 化。null/undefined 渲染为空——Excel 的空单元格。
 */
export function previewCellValue(
  sheet: SheetConfig,
  col: PreviewColumn,
  row: Record<string, unknown>,
  rowIndex: number,
): string {
  if (col.isIndex) return String(indexColumnStart(sheet) + rowIndex);
  const value = row[col.prop];
  if (value === null || value === undefined) return "";
  const numFormat = col.dataStyle?.numFormat;
  if (numFormat !== undefined) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return formatByNumFormat(value, numFormat);
    }
    if (typeof value === "string" && isDateNumFormat(numFormat)) {
      return formatByNumFormat(value, numFormat);
    }
  }
  // 与导出引擎的单元格归一化一致：原始类型直接 String，对象按 JSON 写入
  if (typeof value === "string") return value;
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "bigint"
  )
    return String(value);
  return JSON.stringify(value) ?? "";
}

// ---------------------------------------------------------------------------
// CellStyle -> CSS 映射
// ---------------------------------------------------------------------------

/** Excel 边框档位 -> CSS 宽度/线型段。CSS 没有点划线，dashDot 系降级 dashed。 */
const BORDER_CSS: Record<BorderStyle, string> = {
  hair: "0.5px solid",
  thin: "1px solid",
  medium: "2px solid",
  thick: "3px solid",
  dashed: "1px dashed",
  dotted: "1px dotted",
  double: "3px double",
  mediumDashed: "2px dashed",
  dashDot: "1px dashed",
  mediumDashDot: "2px dashed",
  dashDotDot: "1px dashed",
  mediumDashDotDot: "2px dashed",
  slantDashDot: "2px dashed",
};

/**
 * CellStyle -> CSSProperties。只输出样式声明了的属性，未声明的交给预览表格
 * 的基线（见 styles.demo.tsx 的 PREVIEW_CSS）。视觉近似项：字号 pt 直用（与
 * Excel 同源）；alignment.textRotation 不映射（CSS 旋转按盒子中心而非左下
 * 角锚点，模拟失真大于收益）。
 */
export function cellStyleToCss(style: CellStyle | undefined): CSSProperties {
  if (!style) return {};
  const css: CSSProperties = {};
  if (style.font) {
    const { bold, italic, size, color, name } = style.font;
    if (bold !== undefined) css.fontWeight = bold ? 700 : 400;
    if (italic !== undefined) css.fontStyle = italic ? "italic" : "normal";
    if (size !== undefined) css.fontSize = `${size}pt`;
    if (color !== undefined) css.color = `#${color}`; // 库内 hex 无 #，补上
    if (name !== undefined) css.fontFamily = name;
  }
  if (style.fill?.pattern === "solid" && style.fill.fgColor) {
    css.background = `#${style.fill.fgColor}`;
  }
  if (style.alignment) {
    const { horizontal, vertical, wrapText } = style.alignment;
    if (horizontal) css.textAlign = horizontal;
    if (vertical)
      css.verticalAlign = vertical === "center" ? "middle" : vertical;
    if (wrapText) css.whiteSpace = "pre-wrap";
  }
  if (style.border) {
    const side = (s?: { style: BorderStyle; color?: string }): string | null =>
      s ? `${BORDER_CSS[s.style]} #${s.color ?? "000000"}` : null; // 缺省色 = Excel 自动色（黑）
    const top = side(style.border.top);
    const bottom = side(style.border.bottom);
    const left = side(style.border.left);
    const right = side(style.border.right);
    if (top) css.borderTop = top;
    if (bottom) css.borderBottom = bottom;
    if (left) css.borderLeft = left;
    if (right) css.borderRight = right;
  }
  return css;
}

// ---------------------------------------------------------------------------
// numFormat 预览格式化
// ---------------------------------------------------------------------------

interface Segment {
  literal: boolean; // 双引号字面量段（原样输出）
  text: string; // 已去掉包裹引号
}

/** 按双引号字面量把格式码切段：'yyyy"年"M' -> [code, lit, code]。 */
function splitLiteralSegments(numFormat: string): Segment[] {
  return numFormat
    .split(/("[^"]*")/)
    .filter(Boolean)
    .map((s) => ({
      literal: s.startsWith('"'),
      text: s.startsWith('"') ? s.slice(1, -1) : s,
    }));
}

/**
 * 全部 code 段合并后的文本（格式判别用）。额外剥离 `[...]` 条件/颜色段：
 * 否则 `#,##0;[Red]-#,##0` 里 `[Red]` 的小写 d 会让下面的日期判别误判，
 * 数字列在预览中退化成 String(value)（丢千分位）。
 */
function codeTextOf(numFormat: string): string {
  return splitLiteralSegments(numFormat)
    .filter((s) => !s.literal)
    .map((s) => s.text)
    .join("")
    .replace(/\[[^\]]*\]/g, "");
}

/**
 * 是否日期类格式码。只覆盖演示用到的格式码：含年 token，或含 M/d 且非
 * 百分比（'#,##0.00' / '0.00%' 均不含这些 token，不会误判）。
 */
export function isDateNumFormat(numFormat: string): boolean {
  const code = codeTextOf(numFormat);
  return /y/i.test(code) || (/[Md]/.test(code) && !code.includes("%"));
}

/**
 * 数字格式码渲染：percent = 含 %，decimals = 小数点后 0 的个数，thousands =
 * 数字段含逗号；先舍入后插千分位（与 Excel 一致）。字面量段（如 "¥"）按序
 * 拼在数字前——演示格式码中数字段唯一，% 固定输出在尾部。
 */
function formatNumberByPattern(
  value: number,
  numFormat: string,
  code: string,
): string {
  const percent = code.includes("%");
  const decMatch = code.match(/\.([0#]+)/);
  const decimals = decMatch ? (decMatch[1]!.match(/0/g) ?? []).length : 0;
  const thousands = code.includes(",");
  const fixed = (percent ? value * 100 : value).toFixed(decimals);
  const [intRaw = "", decPart] = fixed.split(".");
  const int = thousands ? intRaw.replace(/\B(?=(\d{3})+$)/g, ",") : intRaw;
  const body = decPart !== undefined ? `${int}.${decPart}` : int;
  const literals = splitLiteralSegments(numFormat)
    .filter((s) => s.literal)
    .map((s) => s.text)
    .join("");
  return percent ? `${literals}${body}%` : `${literals}${body}`;
}

/**
 * 日期格式码渲染：值须是 ISO 日期（或日期时间）字符串——演示数据 orderDate
 * 即此形态；真实导出里 FormatSpec 会把它转成 Excel serial 再交给 numFormat，
 * 两边对同一格式的输出一致。token：yyyy/yy、MM/M、dd/d、HH/H、mm/m——大写
 * M 是月份、小写 m 是分钟（与 Excel 惯例一致，覆盖演示全部格式码）。
 */
function formatDateByTokens(value: string, numFormat: string): string | null {
  const m = value.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/);
  if (!m) return null;
  const y = m[1]!;
  const mo = m[2]!;
  const d = m[3]!;
  const h = m[4] ?? "00";
  const min = m[5] ?? "00";
  return splitLiteralSegments(numFormat)
    .map((seg) => {
      if (seg.literal) return seg.text;
      return seg.text.replace(/yyyy|yy|MM|M|dd|d|HH|H|mm|m/g, (tok) => {
        switch (tok) {
          case "yyyy":
            return y;
          case "yy":
            return y.slice(-2);
          case "MM":
            return mo;
          case "M":
            return String(Number(mo));
          case "dd":
            return d;
          case "d":
            return String(Number(d));
          case "HH":
            return h;
          case "H":
            return String(Number(h));
          case "mm":
            return min;
          case "m":
            return String(Number(min));
          default:
            return tok;
        }
      });
    })
    .join("");
}

/**
 * numFormat 的浏览器预览渲染。只覆盖演示用到的格式码（#,##0.00 / #,##0 /
 * "¥"#,##0.00 / 0.00% / yyyy-MM-dd / yyyy-MM-dd HH:mm / yyyy"年"M"月"d"日"），
 * 其余兜底 String(value)。文本值遇到数字格式码原样返回——Excel 中 numFormat
 * 对文本单元格本就不生效。
 */
export function formatByNumFormat(
  value: number | string,
  numFormat: string,
): string {
  const code = codeTextOf(numFormat);
  if (isDateNumFormat(numFormat)) {
    if (typeof value === "string") {
      const out = formatDateByTokens(value, numFormat);
      if (out !== null) return out;
    }
    return String(value);
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return formatNumberByPattern(value, numFormat, code);
  }
  return String(value);
}
