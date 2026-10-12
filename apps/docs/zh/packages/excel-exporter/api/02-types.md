# API：核心类型

## SheetConfig

| 字段           | 类型                            | 必填 | 说明                                                                                                                                                                 |
| -------------- | ------------------------------- | ---- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `name`         | `string`                        | 是   | 工作表名：非空、≤ 31 字符，不含 `: \ / ? * [ ]`，且不得以单引号开头/结尾。`sheets` 内查重**不区分大小写**（Excel 自身规则——`Sheet1` 与 `SHEET1` 视为同名，会被拒绝） |
| `columns`      | `ColumnConfig[]`                | 是   | 列定义                                                                                                                                                               |
| `data`         | `Record<string, unknown>[]`     | 是   | 行数据                                                                                                                                                               |
| `headerStyle?` | `CellStyle`                     | —    | 表头行默认样式，可被列级 `headerStyle` 覆盖                                                                                                                          |
| `dataStyle?`   | `CellStyle`                     | —    | 全部数据单元格的基底样式，列级 `style` 逐字段深合并覆盖（见样式指南）                                                                                                |
| `indexColumn?` | `boolean \| IndexColumnOptions` | —    | 注入最左侧序号列；`true` 即全部默认值。已有 merges 自动右移一列                                                                                                      |
| `freezeRows?`  | `number`                        | —    | 冻结前 N 行表头；校验为非负整数                                                                                                                                      |
| `merges?`      | `MergeRange[]`                  | —    | 合并单元格（相对数据区定位）                                                                                                                                         |
| `autoFilter?`  | `boolean`                       | —    | 表头自动筛选                                                                                                                                                         |

## ColumnConfig

| 字段           | 类型                     | 必填       | 说明                                                                                       |
| -------------- | ------------------------ | ---------- | ------------------------------------------------------------------------------------------ |
| `prop?`        | `string`                 | 叶子列必填 | 数据行字段名（Element Plus 命名）；分组列（带 `children`）可省略                           |
| `key?`         | `string`                 | —          | `prop` 的废弃别名（2.2 之前命名）；两者同给时 `prop` 优先                                  |
| `label?`       | `string`                 | 是*        | 表头文字（叶子列与分组列都需要）；`label` 或旧名 `header` 至少提供一个（导出时校验）       |
| `header?`      | `string`                 | —          | `label` 的废弃别名（2.2 之前命名）；两者同给时 `label` 优先                                |
| `children?`    | `ColumnConfig[]`         | —          | 分组列：生成多行表头，该表头格自动跨其全部叶子列合并；`children: []` 视为叶子列            |
| `width?`       | `number`                 | —          | 列宽（Excel 字符单位，`0` 表示隐藏列）；校验为有限非负数；仅叶子列生效                     |
| `style?`       | `CellStyle`              | —          | 数据单元格样式（不含表头）；仅叶子列生效                                                   |
| `headerStyle?` | `CellStyle`              | —          | 本列表头样式（含分组表头格），优先于表级 `headerStyle`                                     |
| `format?`      | `FormatSpec \| Function` | —          | 值格式化；仅叶子列；函数在主线程路径执行，浏览器 worker 路径会被剥离（详见 FormatSpec 页） |

带 `children` 的列为分组列，无数据单元格，只贡献表头行。表头行数 = 1 + 列树最大深度（扁平列即 1 行）；叶子列表头纵向跨满剩余表头行，分组列表头横向跨其子树所有叶子列，合并由库自动生成（无需手工写 `merges`）。

## IndexColumnOptions

`SheetConfig.indexColumn` 的选项；简写 `true` 等价于 `{}`。

| 字段           | 类型        | 默认值   | 说明                                                                  |
| -------------- | ----------- | -------- | --------------------------------------------------------------------- |
| `label?`       | `string`    | `"序号"` | 序号列表头文字                                                        |
| `width?`       | `number`    | `6`      | 列宽（Excel 字符单位，`0` 表示隐藏列）                                |
| `start?`       | `number`    | `1`      | 首个数据行显示的序号；第 i 行显示 `start + i`。非负整数（导出时校验） |
| `style?`       | `CellStyle` | —        | 序号列数据单元格样式；与表级 `dataStyle` 的合并规则同普通列           |
| `headerStyle?` | `CellStyle` | —        | 序号列表头样式；优先于表级 `headerStyle`                              |

序号列的值由行号生成、从不读取 `data`；用户列声明保留 prop `__index__` 会被明确报错拒绝。已有 `merges` 自动右移一列，仍指向原目标。

`exportExcel` 会在模式路由之前自动展开该字段（且只展开一次），这正是它「在所有路径生效」的原因。底层入口 `WorkbookBuilder.addSheet()` 与 `exportAsStream()` **不会**自动展开——它们只认识已展开的 `__index__` 列，直连时请先调用 `applyIndexColumn(sheet)`（随包导出，保留字 `INDEX_PROP` 一并导出）。否则 `indexColumn` 会被静默忽略。

## MergeRange

| 字段      | 类型     | 说明                       |
| --------- | -------- | -------------------------- |
| `row`     | `number` | 起始行（0 = 第一条数据行） |
| `col`     | `number` | 起始列（0 = 第一列）       |
| `rowspan` | `number` | 行跨度                     |
| `colspan` | `number` | 列跨度                     |

校验在导出前置执行、所有路径口径一致，报错会指明具体是哪个 merge：四个值必须为整数；`row`/`col` ≥ 0 且 `rowspan`/`colspan` ≥ 1；`col + colspan` 不得超过该表的叶子列数、`row + rowspan` 不得超过数据行数（表头行不在这套坐标系里）；各 range 之间不得重叠（它们整体位于表头区之下，不会与自动生成的表头合并冲突）。`merges` 传非数组同样会被拒绝。

## CellStyle

| 字段         | 类型                                                   | 说明                                                                            |
| ------------ | ------------------------------------------------------ | ------------------------------------------------------------------------------- |
| `font?`      | `{ bold?, italic?, size?, color?, name? }`             | 颜色为 6 位 RGB hex（如 `"FF0000"`）                                            |
| `fill?`      | `{ pattern?: "solid" \| "none", fgColor?, bgColor? }`  | 填充                                                                            |
| `alignment?` | `{ horizontal?, vertical?, wrapText?, textRotation? }` | 对齐（textRotation 0–180）                                                      |
| `border?`    | `{ top?, bottom?, left?, right? }`                     | 边框，每边 `{ style: BorderStyle, color? }`——取值见 [BorderStyle](#borderstyle) |
| `numFormat?` | `string`                                               | Excel 数字格式码                                                                |

样式字段在导出前置校验——非法值以 `{ success: false }` 失败并给出定位到具体表/列的报错，而不是拖到后面变成晦涩的引擎错误或降级导出：颜色必须为 6 位 RGB hex（`"FF0000"`；`"red"` 与 `"#FF0000"` 都会被拒）；`border.*.style` 必须是上文 `BorderStyle` 之一；`font.size` 必须为有限正数；`alignment.textRotation` 必须是 0–180 的**整数**（`45.5` 会被拒绝而非四舍五入）。校验覆盖表级、列级（含分组列被忽略的样式）与 `indexColumn` 样式的全部出现点。

## BorderStyle

每条边框 `style` 的取值——即 Excel 自身的线型词汇，内联自 modern-xlsx 的 `BorderSideData`（使发布的 `.d.ts` 不含依赖导入）：

```ts
type BorderStyle =
  | "thin"
  | "medium"
  | "thick"
  | "dashed"
  | "dotted"
  | "double"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot";
```

每种线型自带粗细——没有单独的宽度参数（`hair` 最细，`thick` 最粗）。预览渲染器对同一批取值做 Excel → CSS 映射，对照表见[预览数据模型参考](/zh/packages/excel-preview/api/02-model#边框-previewborder)。

## ExportMode / ExportPhase

```ts
type ExportMode = "auto" | "main" | "worker" | "stream";
type ExportPhase = "init" | "build" | "download";
```

## 完整导入

```ts
import type {
  SheetConfig,
  ColumnConfig,
  CellStyle,
  MergeRange,
  FormatSpec,
  ExportOptions,
  ExportResult,
  ExportMode,
  ExportPhase,
  BorderStyle,
} from "@marcusok/excel-exporter";
```
