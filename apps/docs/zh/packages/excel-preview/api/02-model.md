# 数据模型（PreviewWorkbook）

`parseWorkbookBytes`（与 Worker）产出纯 JSON 模型——结构化克隆安全、框架无关，与内置渲染器消费的数据完全一致。

## 结构

```ts
interface PreviewWorkbook {
  sheets: PreviewSheet[];
  activeSheetIndex: number; // 0 起，来自文件 activeTab
  dateSystem: "date1900" | "date1904"; // 决定序列号 → 日期的换算
}

interface PreviewSheet {
  name: string;
  visible: boolean; // hidden / veryHidden 保留在列表里
  showGridLines: boolean;
  rightToLeft: boolean;
  rowCount: number; // 1 起数据边界
  colCount: number;
  rows: PreviewRow[]; // 稀疏：无内容/行高/隐藏标记的行缺席
  colSpans: PreviewColSpan[]; // <col> 跨度（width/hidden/customWidth），min/max 1 起
  merges: PreviewMerge[]; // 0 起 { row, col, rowSpan, colSpan }
  frozenRows: number; // 0 = 无
  frozenCols: number;
  styles: PreviewStyles; // 全工作簿共享（styles.xml 级）
}

interface PreviewRow {
  index: number; // 1 起
  height: number | null; // pt；null → 默认（15pt）
  hidden: boolean;
  cells: PreviewCell[];
}

interface PreviewCell {
  col: number; // 0 起
  type: "number" | "string" | "boolean" | "error" | "formulaStr";
  value: string | null; // 数字保持字符串（"45678.5"）；null = 空
  styleIndex: number | null; // styles.xfs 下标；null = 默认样式
}

interface PreviewStyles {
  fonts: PreviewFont[]; // { name, size(pt), bold, italic, underline, strike, color }
  fills: PreviewFill[]; // none | solid | pattern | gradient（stops 已解析）
  borders: PreviewBorder[]; // 每边 { style, color } | null
  xfs: PreviewXf[]; // { fontId, fillId, borderId, numFmtCode, alignment }
}
```

消费方要点：

- **值是原始字符串。** 数字保持全精度字符串形态；需要数值时自行 `Number(value)`。渲染器的格式化路径就是这么做的。
- **`numFmtCode` 已预解析。** 内置 id 已过 Excel 行为表（14 → `m/d/yyyy`），自定义 id 来自文件 `numFmts`。永远不会是 `undefined`——下限是 `"General"`。
- **颜色是 CSS 字符串**（`#rrggbb` / `rgba(...)`），已经过主题/indexed 覆盖层解析；`null` 表示自动（渲染默认：黑字、无填充）。
- **`formulaStr` 单元格**是公式的字符串缓存值；数字缓存的公式单元格以 `type: "number"` 携带缓存值出现（与 Excel 显示一致）。
- **空单元格缺席**——模型两个维度都是稀疏的；渲染方应把"缺席"当作默认样式的空格。

## 自行格式化单元格

渲染器使用的格式化器原样导出，供自研渲染器复用：

```ts
import { formatCellValue } from "@marcusok/excel-preview";

const { text, color } = formatCellValue(
  cell.type,
  cell.value,
  sheet.styles.xfs[cell.styleIndex ?? 0]?.numFmtCode ?? "General",
  workbook.dateSystem,
);
// text: 显示文本；color: 格式含 [Red] 类颜色段时的 "#ff0000"
```
