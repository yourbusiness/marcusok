# API：核心类型

## SheetConfig

| 字段           | 类型                            | 必填 | 说明                                                                     |
| -------------- | ------------------------------- | ---- | ------------------------------------------------------------------------ |
| `name`         | `string`                        | 是   | 工作表名：非空、≤ 31 字符，不含 `: \ / ? * [ ]`，且不得以单引号开头/结尾 |
| `columns`      | `ColumnConfig[]`                | 是   | 列定义                                                                   |
| `data`         | `Record<string, unknown>[]`     | 是   | 行数据                                                                   |
| `headerStyle?` | `CellStyle`                     | —    | 表头行默认样式，可被列级 `headerStyle` 覆盖                              |
| `dataStyle?`   | `CellStyle`                     | —    | 全部数据单元格的基底样式，列级 `style` 逐字段深合并覆盖（见样式指南）    |
| `indexColumn?` | `boolean \| IndexColumnOptions` | —    | 注入最左侧序号列；`true` 即全部默认值。已有 merges 自动右移一列          |
| `freezeRows?`  | `number`                        | —    | 冻结前 N 行表头；校验为非负整数                                          |
| `merges?`      | `MergeRange[]`                  | —    | 合并单元格（相对数据区定位）                                             |
| `autoFilter?`  | `boolean`                       | —    | 表头自动筛选                                                             |

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

带 `children` 的列为分组列，无数据单元格，只贡献表头行。表头行数 = 列树最大深度；叶子列表头纵向跨满剩余表头行，分组列表头横向跨其子树所有叶子列，合并由库自动生成（无需手工写 `merges`）。

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

## MergeRange

| 字段      | 类型     | 说明                       |
| --------- | -------- | -------------------------- |
| `row`     | `number` | 起始行（0 = 第一条数据行） |
| `col`     | `number` | 起始列（0 = 第一列）       |
| `rowspan` | `number` | 行跨度                     |
| `colspan` | `number` | 列跨度                     |

## CellStyle

| 字段         | 类型                                                   | 说明                                 |
| ------------ | ------------------------------------------------------ | ------------------------------------ |
| `font?`      | `{ bold?, italic?, size?, color?, name? }`             | 颜色为 6 位 RGB hex（如 `"FF0000"`） |
| `fill?`      | `{ pattern?: "solid" \| "none", fgColor?, bgColor? }`  | 填充                                 |
| `alignment?` | `{ horizontal?, vertical?, wrapText?, textRotation? }` | 对齐（textRotation 0–180）           |
| `border?`    | `{ top?, bottom?, left?, right? }`                     | 边框，每边 `{ style, color? }`       |
| `numFormat?` | `string`                                               | Excel 数字格式码                     |

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
