import type { AlignmentData, BorderStyle } from "@marcusok/xlsx-core";

/**
 * @marcusok/excel-preview 公共类型。
 *
 * 预览数据模型（PreviewWorkbook 一族）是纯 JSON：解析（通常在 Worker 内）
 * 产出后经结构化克隆传回主线程，渲染层只消费该模型。所有字段对缺失形态
 * （null/undefined）保持防御——modern-xlsx 读取侧的"无内容"形态不统一
 * （实测 null / 空数组 / undefined 三种并存）。
 */

/** 解析后的工作簿（纯数据，结构化克隆安全）。 */
export interface PreviewWorkbook {
  sheets: PreviewSheet[];
  /** 打开时选中的 sheet（0-based；解析层已对越界 activeTab 钳制到 0、对
   * 指向隐藏页签的 activeTab 顺延到首个可见 sheet，渲染层的越界回退是
   * 第二重兜底，见 render() 的初始 sheet 处理）。 */
  activeSheetIndex: number;
  /** 日期系统：决定序列号 → 日期的偏移（1904 系统需 +1462 天）。 */
  dateSystem: "date1900" | "date1904";
}

export interface PreviewSheet {
  name: string;
  /** false = 隐藏（hidden / veryHidden 都不出现在页签里）。 */
  visible: boolean;
  /** 文件声明的视图开关（缺省 true）。 */
  showGridLines: boolean;
  rightToLeft: boolean;
  /** 数据边界（1-based，含样式格；合并区末行/列已并入——纯合并不写覆盖格，
   * 只按内容归约会把合并区截小，见 parse/model.ts）。 */
  rowCount: number;
  colCount: number;
  /** 稀疏行（按 index 升序，无内容行缺席）。 */
  rows: PreviewRow[];
  /** 列跨度定义（1-based min/max，与 styles.xml 的 <col> 对应）。 */
  colSpans: PreviewColSpan[];
  /** 合并单元格（0-based 行列）。 */
  merges: PreviewMerge[];
  /** 冻结行数 / 列数（0 = 无）。 */
  frozenRows: number;
  frozenCols: number;
  /** 解析合并后的样式表（含主题色/indexed 覆盖层结果）。 */
  styles: PreviewStyles;
}

export interface PreviewRow {
  /** 1-based 行号。 */
  index: number;
  /** 行高（pt）；null/缺省 = 默认行高（15pt）。 */
  height: number | null;
  hidden: boolean;
  cells: PreviewCell[];
}

export interface PreviewCell {
  /** 0-based 列号。 */
  col: number;
  /** 归一化后的单元格类型（number/string/boolean/error + 2 种字符串来源）。 */
  type: PreviewCellType;
  /** 原始值字符串形态（数字为字符串，如 "45678.5"；空单元格无该字段）。 */
  value: string | null;
  /** 样式索引（styles.xfs 下标；null = 默认样式）。 */
  styleIndex: number | null;
}

export type PreviewCellType =
  "number" | "string" | "boolean" | "error" | "formulaStr";

export interface PreviewColSpan {
  min: number;
  max: number;
  /** Excel 字符宽度单位；缺省 8.43。 */
  width: number;
  hidden: boolean;
  /** 文件声明的 customWidth 标志（元信息）。渲染层不依赖它判定列宽：
   * <col> 出现即采用其 width——引擎对未写 width 的 col 兜底 8.43（与默认
   * 一致），只写 width 不写标志的第三方产物列宽不能丢（见 buildLayout）。 */
  customWidth: boolean;
}

export interface PreviewMerge {
  /** 0-based，主单元格位置。 */
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
}

/** 解析合并后的样式表（modern-xlsx 解析结果 + 颜色覆盖层，按数组下标 1:1）。 */
export interface PreviewStyles {
  fonts: PreviewFont[];
  fills: PreviewFill[];
  borders: PreviewBorder[];
  /** xf：cellXfs 逐项 + 已解析的格式码（内置 id 经 Excel 行为表转换）。 */
  xfs: PreviewXf[];
}

export interface PreviewFont {
  name: string | null;
  /** pt；null = 默认 11。 */
  size: number | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  /** "#RRGGBB"；null = 自动（渲染默认黑）。 */
  color: string | null;
}

export type PreviewFill =
  | { kind: "none" }
  | { kind: "solid"; fgColor: string | null }
  | { kind: "pattern"; fgColor: string | null; bgColor: string | null }
  | {
      kind: "gradient";
      degree: number;
      stops: { position: number; color: string }[];
    };

export interface PreviewBorderSide {
  style: BorderStyle;
  color: string | null;
}

export interface PreviewBorder {
  left: PreviewBorderSide | null;
  right: PreviewBorderSide | null;
  top: PreviewBorderSide | null;
  bottom: PreviewBorderSide | null;
  diagonal: PreviewBorderSide | null;
  /**
   * 对角线方向：true = "/"（diagonalUp）；缺省/false = "\"（diagonalDown，
   * Excel 常见形态）。仅有 diagonal 样式而 Up/Down 标志全缺的文件不显示
   * 对角线（Excel 语义，解析层已置 diagonal 为 null）。
   */
  diagonalUp?: boolean;
}

export interface PreviewXf {
  fontId: number;
  fillId: number;
  borderId: number;
  /**
   * 已解析的数字格式码（内置 id → Excel 实际行为串；自定义 → 原码）。
   * "General" 表示无格式；null 不出现（解析层兜底）。
   */
  numFmtCode: string;
  alignment: AlignmentData | null;
}

// ---------------------------------------------------------------------------
// 高层 API（createPreview）
// ---------------------------------------------------------------------------

/** 输入源：文件对象或原始字节。 */
export type PreviewSource = File | Blob | Uint8Array | ArrayBuffer;

export interface PreviewOptions {
  source: PreviewSource;
  /** 加密文件的密码（modern-xlsx Agile AES-256）。 */
  password?: string;
  /** 初始 sheet（名称或 0-based 索引）；缺省 = 文件的 activeTab。名称不存在
   * 或索引越界时静默回退 activeTab（不触发 onError——对照实例方法 setSheet
   * 对无效入参会回调报错：初始选项容错、运行期切换严格，两者语义有意不同）。 */
  sheet?: string | number;
  /** 表头（行号列/列标头）显示开关；缺省 true。表头是预览自身 UI，无文件级
   * 声明可遵循（对照：showGridLines 确实读取文件的视图开关）。 */
  showHeaders?: boolean;
  /** 覆盖文件声明的网格线开关；缺省遵循文件（默认显示）。 */
  showGridLines?: boolean;
  /** 是否渲染 sheet 页签栏；默认 true。 */
  showTabs?: boolean;
  /** 渲染就绪回调：首次解析渲染完成后与每次 sheet 切换完成后都会触发
   * （不是仅触发一次；duration.parse 复用首次解析耗时）。 */
  onParsed?: (info: PreviewParsedInfo) => void;
  /** 错误回调（解析/渲染致命错误）。 */
  onError?: (error: PreviewError) => void;
}

export interface PreviewParsedInfo {
  sheetNames: string[];
  sheetCount: number;
  /** 当前 sheet 的行列规模。 */
  rowCount: number;
  colCount: number;
  /** 各阶段耗时（毫秒）。 */
  duration: { parse: number; render: number; total: number };
}

export type PreviewErrorCode =
  | "PASSWORD_PROTECTED"
  | "LEGACY_FORMAT"
  | "UNSUPPORTED"
  | "CORRUPT"
  | "WASM"
  | "UNKNOWN";

export interface PreviewError {
  code: PreviewErrorCode;
  message: string;
  /** 原始错误（诊断用）。 */
  cause?: unknown;
}

export interface PreviewInstance {
  /** 卸载渲染 DOM 并释放资源（Worker 为模块级共享，不随实例销毁）。 */
  destroy(): void;
  /** 切换 sheet（名称或 0-based 索引）。 */
  setSheet(nameOrIndex: string | number): void;
  /** 全部 sheet 名（含隐藏页签，按文件顺序）。 */
  getSheetNames(): string[];
}
