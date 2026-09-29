/**
 * 预览模型构建：modern-xlsx Workbook → PreviewWorkbook（纯 JSON）。
 *
 * 该层吞掉读取侧所有"无内容形态不统一"问题（实测 null / 空数组 / undefined
 * 三种并存：mergeCells/frozenPane/columns 为 undefined、splitPane/view/
 * tabColor 为 null、hyperlinks/comments 为 []），输出统一形态；并把样式链
 * （fonts/fills/borders/cellXfs/numFmts）解析成渲染层直接可用的扁平结构，
 * 颜色经覆盖层（styles-overlay）找回主题色/indexed。
 */
import type {
  CellData,
  BorderStyle,
  FillData,
  RowData,
  ColumnInfo,
  Worksheet,
  Workbook,
} from "@marcusok/xlsx-core";
import { getBuiltinFormat } from "@marcusok/xlsx-core";
import { resolveFormatCode } from "../numfmt/builtin-table";
import { argbToCss, type ThemePalette } from "../color";
import {
  buildStylesOverlay,
  resolveColorSpec,
  type OverlayBorder,
  type OverlayFill,
  type OverlayFont,
  type StylesOverlay,
} from "./styles-overlay";
import type {
  PreviewBorder,
  PreviewCell,
  PreviewCellType,
  PreviewColSpan,
  PreviewFill,
  PreviewFont,
  PreviewRow,
  PreviewSheet,
  PreviewStyles,
  PreviewWorkbook,
} from "../types";

/** 归一化 cellType：字符串类（sharedString/inlineStr）合并，stub 视为空。 */
function normalizeCellType(t: CellData["cellType"]): PreviewCellType {
  switch (t) {
    case "number":
      return "number";
    case "boolean":
      return "boolean";
    case "error":
      return "error";
    case "formulaStr":
      return "formulaStr";
    case "stub":
    case "sharedString":
    case "inlineStr":
    default:
      return "string";
  }
}

function decodeRef(reference: string): number {
  // 快路径：直接从 A1 引用解析列号（0-based）。decodeCellRef 往返一次，
  // 百万格级模型构建下字符串解析成本可见，这里用本地实现。
  let col = 0;
  let i = 0;
  for (; i < reference.length; i++) {
    const c = reference.charCodeAt(i);
    if (c >= 65 && c <= 90) col = col * 26 + (c - 64);
    else if (c >= 97 && c <= 122) col = col * 26 + (c - 96);
    else break;
  }
  return col - 1;
}

/** 空值/undefined 归一化到统一缺省形态的工具。 */
function num(v: number | null | undefined, fallback: number): number {
  return typeof v === "number" && Number.isFinite(v) ? v : fallback;
}

/** ECMA-376 ST_BorderStyle 合法值（border style 属性白名单）。 */
const BORDER_STYLES = new Set([
  "thin",
  "medium",
  "thick",
  "double",
  "dashed",
  "dotted",
  "hair",
  "dashDot",
  "dashDotDot",
  "mediumDashDot",
  "mediumDashDotDot",
  "mediumDashed",
  "slantDashDot",
]);

function borderStyleOf(raw: string): BorderStyle {
  // 未知值（方言文件）按 thin 近似：保留"此处有边框"的视觉提示，与
  // css.ts 对未知样式的兜底口径一致（"none" 在 borderFromOverlay 已拦截）
  return BORDER_STYLES.has(raw) ? (raw as BorderStyle) : "thin";
}

/** 覆盖层重建路径：styles.xml 文档序 1:1（含自闭合空项），与 cellXfs 的 id 对齐。 */
function fontFromOverlay(f: OverlayFont, palette: ThemePalette): PreviewFont {
  return {
    name: f.name,
    size: f.size,
    bold: f.bold,
    italic: f.italic,
    underline: f.underline,
    strike: f.strike,
    color: resolveColorSpec(f.color, palette),
  };
}

function fillFromOverlay(f: OverlayFill, palette: ThemePalette): PreviewFill {
  const stops = f.gradient?.stops;
  if (f.gradient && stops && stops.length > 0) {
    return {
      kind: "gradient",
      degree: f.gradient.degree,
      stops: stops
        .map((s, si) => ({
          position: num(s.position, si / Math.max(1, stops.length - 1)),
          color: resolveColorSpec(s.spec, palette) ?? "#000000",
        }))
        .sort((a, b) => a.position - b.position),
    };
  }
  const pattern = f.patternType ?? "";
  if (pattern === "" || pattern === "none") return { kind: "none" };
  if (pattern === "solid") {
    return {
      kind: "solid",
      fgColor: resolveColorSpec(f.fgColor, palette),
    };
  }
  return {
    kind: "pattern",
    fgColor: resolveColorSpec(f.fgColor, palette),
    bgColor: resolveColorSpec(f.bgColor, palette),
  };
}

function borderFromOverlay(
  b: OverlayBorder,
  palette: ThemePalette,
): PreviewBorder {
  const side = (s: OverlayBorder["left"]) =>
    s && s.style && s.style !== "none"
      ? {
          style: borderStyleOf(s.style),
          color: resolveColorSpec(s.color, palette),
        }
      : null;
  return {
    left: side(b.left),
    right: side(b.right),
    top: side(b.top),
    bottom: side(b.bottom),
    diagonal: side(b.diagonal),
  };
}

/**
 * 引擎降级路径：overlay 不可用（styles.xml 缺失/zip 损坏）时兜底。
 * 引擎只回显式 RGB（theme/indexed 被读取侧丢弃），且其数组会跳过自闭合
 * 元素（可能错位）——两条缺陷在覆盖层缺席时无从补救，按原样兜底。
 */
function stylesFromEngine(
  wb: Workbook,
): Pick<PreviewStyles, "fonts" | "fills" | "borders"> {
  const st = wb.styles;
  const fonts: PreviewFont[] = (st.fonts ?? []).map((f) => ({
    name: f?.name ?? null,
    size: typeof f?.size === "number" ? f.size : null,
    bold: !!f?.bold,
    italic: !!f?.italic,
    underline: !!f?.underline,
    strike: !!f?.strike,
    color: argbToCss(f?.color),
  }));

  const fills: PreviewFill[] = (st.fills ?? []).map((f: FillData) => {
    // GradientFillData.stops 的元素在类型层收敛为 any（渐变 stop 形态来自
    // 引擎声明文件），本地收窄保证字段访问的类型安全
    interface GradientStopLike {
      position?: number | null;
      color?: string | null;
    }
    const gradient = f?.gradientFill as
      { degree?: number | null; stops?: GradientStopLike[] } | undefined;
    const stopsList = gradient?.stops;
    if (gradient && Array.isArray(stopsList) && stopsList.length > 0) {
      return {
        kind: "gradient",
        degree: num(gradient.degree, 90),
        stops: stopsList
          .map((s, si) => ({
            position: num(s.position, si / Math.max(1, stopsList.length - 1)),
            color: argbToCss(s.color) ?? "#000000",
          }))
          .sort((a, b) => a.position - b.position),
      };
    }
    const pattern = f?.patternType ?? "";
    if (pattern === "" || pattern === "none") return { kind: "none" };
    if (pattern === "solid") {
      return { kind: "solid", fgColor: argbToCss(f.fgColor) };
    }
    return {
      kind: "pattern",
      fgColor: argbToCss(f.fgColor),
      bgColor: argbToCss(f.bgColor),
    };
  });

  const borders = (st.borders ?? []).map((b) => {
    const side = (side: "left" | "right" | "top" | "bottom" | "diagonal") => {
      const s = b?.[side];
      if (!s || !s.style) return null;
      return { style: s.style, color: argbToCss(s.color) };
    };
    return {
      left: side("left"),
      right: side("right"),
      top: side("top"),
      bottom: side("bottom"),
      diagonal: side("diagonal"),
    };
  });

  return { fonts, fills, borders };
}

function buildStyles(
  wb: Workbook,
  overlay: StylesOverlay | null,
  palette: ThemePalette,
): PreviewStyles {
  const st = wb.styles;

  // fonts/fills/borders 优先走覆盖层重建（styles.xml 文档序、含自闭合空项，
  // 与 cellXfs 的 fontId/fillId/borderId 严格对齐）：引擎解析跳过自闭合
  // 元素（<border/> 等，真实 Excel 产物的 border#0 常态）导致其数组整体
  // 前移错位——实测 borderId=0 的默认单元格会拿到文件 border#1 的样式、
  // 末位引用越界丢失。cellXfs 本身无自闭合问题（xf 的属性在 start 事件
  // 即读完），保持引擎来源。
  const base = overlay
    ? {
        fonts: overlay.fonts.map((f) => fontFromOverlay(f, palette)),
        fills: overlay.fills.map((f) => fillFromOverlay(f, palette)),
        borders: overlay.borders.map((b) => borderFromOverlay(b, palette)),
      }
    : stylesFromEngine(wb);

  // 自定义 numFmts：数组形态 [{id, formatCode}] → Map（实测数组非映射）
  const customFmt = new Map<number, string>();
  for (const nf of st.numFmts ?? []) {
    if (nf && typeof nf.id === "number" && nf.formatCode) {
      customFmt.set(nf.id, nf.formatCode);
    }
  }

  const xfs = (st.cellXfs ?? []).map((xf) => ({
    fontId: num(xf?.fontId, 0),
    fillId: num(xf?.fillId, 0),
    borderId: num(xf?.borderId, 0),
    numFmtCode: resolveFormatCode(num(xf?.numFmtId, 0), customFmt, (id) =>
      getBuiltinFormat(id),
    ),
    alignment: xf?.alignment ?? null,
  }));

  return { ...base, xfs };
}
function buildSheet(ws: Worksheet, styles: PreviewStyles): PreviewSheet {
  const rows: PreviewRow[] = [];
  let maxCol = 0;
  for (const r of (ws.rows ?? []) as RowData[]) {
    if (!r || typeof r.index !== "number") continue;
    const cells: PreviewCell[] = [];
    for (const c of r.cells ?? []) {
      if (!c || typeof c.reference !== "string") continue;
      const col = decodeRef(c.reference);
      if (col < 0) continue;
      // 空单元格（无值且无样式）不进模型——稀疏化，虚拟渲染受益
      if (c.value == null && c.styleIndex == null) continue;
      cells.push({
        col,
        type: normalizeCellType(c.cellType),
        value: c.value == null ? null : String(c.value),
        styleIndex: typeof c.styleIndex === "number" ? c.styleIndex : null,
      });
      if (col + 1 > maxCol) maxCol = col + 1;
    }
    // 无内容空行跳过（隐藏行保留）；自定义行高的空行必须保留——引擎会
    // 输出 cells 为空但带 height 的行（实测 <row ht customHeight> 无 cell 的
    // 形态），跳过会让渲染层把该行回落默认行高，拉高留白的行整体塌掉
    if (cells.length === 0 && !r.hidden && r.height == null) continue;
    rows.push({
      index: r.index,
      height: typeof r.height === "number" ? r.height : null,
      hidden: !!r.hidden,
      cells,
    });
  }
  rows.sort((a, b) => a.index - b.index);
  let rowCount = rows.length > 0 ? rows[rows.length - 1].index : 0;

  const colSpans: PreviewColSpan[] = ((ws.columns ?? []) as ColumnInfo[]).map(
    (c) => ({
      min: num(c?.min, 1),
      max: num(c?.max, c?.min ?? 1),
      width: num(c?.width, 8.43),
      hidden: !!c?.hidden,
      customWidth: !!c?.customWidth,
    }),
  );
  for (const s of colSpans) {
    if (s.max > maxCol) maxCol = s.max;
  }

  const merges = (ws.mergeCells ?? [])
    .map((ref) => {
      // "A1:C3" → 0-based {row, col, rowSpan, colSpan}
      const m = /^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/.exec(ref ?? "");
      if (!m) return null;
      const c1 = decodeRef(m[1]);
      const c2 = decodeRef(m[3]);
      const r1 = Number(m[2]) - 1;
      const r2 = Number(m[4]) - 1;
      return {
        row: Math.min(r1, r2),
        col: Math.min(c1, c2),
        rowSpan: Math.abs(r2 - r1) + 1,
        colSpan: Math.abs(c2 - c1) + 1,
      };
    })
    .filter((x): x is NonNullable<typeof x> => x !== null);

  // 数据边界并入合并区末行/列：Excel 对纯合并不写覆盖格（实测引擎 rows 只
  // 回主格行、无自定义行高/列宽的空行空列整体缺席），只按内容归约时
  // "A1:C3 合并、仅 A1 有值"的 rowCount/colCount 会是 1×1，渲染层把合并
  // 尺寸截断到 1×1（Excel 应显示 3×3）
  for (const m of merges) {
    if (m.row + m.rowSpan > rowCount) rowCount = m.row + m.rowSpan;
    if (m.col + m.colSpan > maxCol) maxCol = m.col + m.colSpan;
  }

  const frozen = ws.frozenPane;
  const view = ws.view;

  return {
    name: ws.name ?? "",
    // state getter：visible/hidden/veryHidden（隐藏页签不渲染、不参与切换）
    visible: ws.state === "visible" || ws.state == null,
    showGridLines: view?.showGridLines !== false,
    rightToLeft: !!view?.rightToLeft,
    rowCount: Math.max(rowCount, 0),
    colCount: Math.max(maxCol, 0),
    rows,
    colSpans,
    merges,
    frozenRows: typeof frozen?.rows === "number" ? Math.max(0, frozen.rows) : 0,
    frozenCols: typeof frozen?.cols === "number" ? Math.max(0, frozen.cols) : 0,
    styles,
  };
}

/**
 * 构建 PreviewWorkbook。在 Worker 或主线程均可调用；先做覆盖层（只解压
 * styles/theme 两个小部件），再逐 sheet 归一化。
 */
export function buildPreviewWorkbook(
  wb: Workbook,
  bytes: Uint8Array,
): PreviewWorkbook {
  const { overlay, palette } = buildStylesOverlay(bytes);
  const styles = buildStyles(wb, overlay, palette);

  const names = wb.sheetNames ?? [];
  const sheets: PreviewSheet[] = [];
  for (const name of names) {
    const ws = wb.getSheet(name);
    if (!ws) {
      // 名字取不到（异常文件）：占位空 sheet，保持索引对齐
      sheets.push({
        name,
        visible: false,
        showGridLines: true,
        rightToLeft: false,
        rowCount: 0,
        colCount: 0,
        rows: [],
        colSpans: [],
        merges: [],
        frozenRows: 0,
        frozenCols: 0,
        styles,
      });
      continue;
    }
    sheets.push(buildSheet(ws, styles));
  }

  // activeTab：bookViews[0].activeTab（0-based）；越界由渲染层钳制
  const activeTabRaw = wb.workbookViews?.[0]?.activeTab;
  const activeSheetIndex =
    typeof activeTabRaw === "number" &&
    activeTabRaw >= 0 &&
    activeTabRaw < sheets.length
      ? activeTabRaw
      : 0;

  return {
    sheets,
    activeSheetIndex,
    dateSystem: wb.dateSystem === "date1904" ? "date1904" : "date1900",
  };
}
