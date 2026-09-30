# API：表格与 ECharts 便捷入口

两个便捷入口：直接吃你手上已有的数据形态——`columns + data` 表格对，或 ECharts option 的一个小型结构子集——把它们归一成 [`SheetConfig`](/zh/packages/excel-exporter/api/02-types#sheetconfig)，再交给 [`exportExcel`](/zh/packages/excel-exporter/api/01-export-excel)。两者都只是薄适配层：它们能表达的一切，用 `exportExcel` 直接写同样能表达。

## 与 exportExcel 的关系

`exportTable` 与 `exportEcharts` 是对 `exportExcel` 的两步封装：

1. 纯转换（`tableExportToOptions` / `echartsExportToOptions`）把便捷选项变成一份 `ExportOptions`；
2. 把结果交给 `exportExcel`。

由此带来几点值得注意的结论：

- 这两个类型确实声明的公共 [`ExportOptions`](/zh/packages/excel-exporter/api/01-export-excel#exportoptions) 字段——`filename`、`mode`、`onProgress`、`onPhase`、`download`——语义与主入口完全一致（同样的路由、同样的进度契约、同样的 `{ success, blob, engine, mode, duration, rowCount, error }` 结果）。
- `ExportOptions` 的 `overlay` 字段**不在** `TableExportOptions` / `EChartsExportOptions` 里，两个转换函数也不透传它。因此遮罩仍按默认行为（开启）运行，但无法通过这两个封装自定义或关闭——确实需要时，请自行转换后调用 `exportExcel`。
- 两个入口都是单 sheet：各自产出的 `sheets` 数组有且只有一项。

```ts
import {
  exportTable,
  exportEcharts,
  tableToSheet,
  tableExportToOptions,
  echartsToSheet,
  echartsExportToOptions,
} from "@marcusok/excel-exporter";
```

它们没有单独的子路径入口——主入口整体重导出了这两个适配层（`src/table-export.ts`、`src/echarts-export.ts`），函数与类型一并在内。

## exportTable

```ts
exportTable(options: TableExportOptions): Promise<ExportResult>
```

由 `columns + data` 导出一张 sheet，列描述符沿用 Element Plus / Ant Design 组件已在用的那些字段名。sheet 名默认 `"Sheet1"`，`freezeRows` / `autoFilter` / `merges` / `dataStyle` / `indexColumn` 原样透传给 sheet。

### TableExportOptions

`TableExportOptions` 在 `TableSheetInput` 之上追加了导出级字段：

| 字段          | 类型                                               | 必填 | 说明                                                      |
| ------------- | -------------------------------------------------- | ---- | --------------------------------------------------------- |
| `columns`     | `TableColumnInput[]`                               | 是   | 列描述符，会被归一为 `ColumnConfig`（见下）               |
| `data`        | `Record<string, unknown>[]`                        | 是   | 行对象，以解析出的 `prop` 为键                            |
| `filename`    | `string`                                           | 是   | 下载文件名，不以 `.xlsx` 结尾时末尾自动追加               |
| `sheetName`   | `string`                                           | —    | 默认 `"Sheet1"`                                           |
| `freezeRows`  | `number`                                           | —    | 冻结的表头行数，校验为非负整数                            |
| `autoFilter`  | `boolean`                                          | —    | 表头筛选，范围覆盖最后一行表头加全部数据行                |
| `merges`      | `MergeRange[]`                                     | —    | 合并单元格范围，相对数据区域                              |
| `dataStyle`   | `CellStyle`                                        | —    | 全部数据单元格的基底样式，列级 `style` 在其上逐字段深合并 |
| `indexColumn` | `boolean \| IndexColumnOptions`                    | —    | 注入最左侧序号列，`true` 表示全默认；`merges` 会自动右移  |
| `mode`        | `"auto" \| "main" \| "worker" \| "stream"`         | —    | 默认 `"auto"`                                             |
| `onProgress`  | `(progress: number) => void`                       | —    | 0 → 1，所有路径上首尾 0 与 1 各上报一次                   |
| `onPhase`     | `(phase: ExportPhase, durationMs: number) => void` | —    | `init` / `build` / `download` 阶段耗时                    |
| `download`    | `boolean`                                          | —    | 默认 `true`；`false` 只返回 Blob                          |

### TableColumnInput

刻意做到零依赖的列描述符——签名里不会渗入 Element Plus 或 Ant Design 的类型。

| 字段           | 类型                     | 必填   | 说明                                                                                    |
| -------------- | ------------------------ | ------ | --------------------------------------------------------------------------------------- |
| `prop?`        | `string`                 | 叶子列 | 数据行上的字段名（Element Plus / 本库命名）；优先于 `key` 与 `dataIndex`                |
| `key?`         | `string`                 | —      | `prop` 的旧名（2.2 之前的命名），`prop` 缺失时启用                                      |
| `dataIndex?`   | `string`                 | —      | `prop` 的 Ant Design 别名，兜底项                                                       |
| `label?`       | `string \| number`       | 是     | 表头文案；优先于 `header` 与 `title`。有限数字会被转成字符串                            |
| `header?`      | `string \| number`       | —      | `label` 的旧名（2.2 之前的命名）                                                        |
| `title?`       | `string \| number`       | —      | `label` 的 Ant Design 别名，兜底项                                                      |
| `width?`       | `number`                 | —      | 列宽（Excel 字符单位，`0` 表示隐藏该列）；仅叶子列                                      |
| `style?`       | `CellStyle`              | —      | 数据单元格样式，在表级 `dataStyle` 之上深合并；仅叶子列                                 |
| `headerStyle?` | `CellStyle`              | —      | 本列的表头样式（含分组表头格）；叶子列与分组列都支持                                    |
| `format?`      | `FormatSpec \| Function` | —      | 取值格式化（见 [FormatSpec](/zh/packages/excel-exporter/api/03-format-spec)）；仅叶子列 |
| `children?`    | `TableColumnInput[]`     | —      | 分组列：生成多行表头并递归。`children: []` 视作叶子列                                   |

别名解析就是一串 `??`，因此优先级为：字段名 `prop` → `key` → `dataIndex`；表头文案 `label` → `header` → `title`。

`children` 为非空数组的列是**分组列**：它产生一个横跨整个子树的表头格，且只要求有表头文案——`prop`、`width`、`style`、`format` 对它没有意义（没有数据单元格），会被丢弃。叶子表头纵向铺满剩余表头行，分组表头横向铺满自己的叶子子树。表头的合并区域自动生成，与[核心类型](/zh/packages/excel-exporter/api/02-types#columnconfig)页对 `ColumnConfig` 的描述一致。

校验发生在转换期，立即失败并给出具体文案：

| 情形                       | 错误文案                                                                                                   |
| -------------------------- | ---------------------------------------------------------------------------------------------------------- |
| 叶子/分组列缺可用表头文案  | `[excel-exporter] table column "<prop 或 group-N>" has no usable header. Provide label, header, or title.` |
| 叶子列有表头但缺可用字段名 | `[excel-exporter] table column #<index> has no usable prop. Provide prop, key, or dataIndex.`              |
| `children` 构成环          | `[excel-exporter] circular children reference in table columns`                                            |

几个实操中会碰到的细节：

- 表头先于字段名检查，所以两者都缺的列先报表头错误。该文案里的 `<prop 或 group-N>` 在有解析出的 `prop` 时就是它，否则是占位符 `group-<index>`（首个顶层列为 `group-0`）。
- `<index>` 是该列在**同层兄弟**中的位置——坏掉的子列报的是它在 `children` 里的下标，不是绝对位置。
- 空串不算表头：表头必须是非空字符串，或有限数字（会被转成字符串）。`label: ""` 会抛错。
- 环检测之所以存在，是因为这个转换跑在 `exportExcel` 拍平列树**之前**：没有它，循环 `children` 会表现为栈溢出，而不是一条可读的错误。

### TableSheetInput → SheetConfig

映射由 `tableToSheet` 完成，几乎是逐一对应：

| `TableSheetInput` | `SheetConfig` | 说明                                |
| ----------------- | ------------- | ----------------------------------- |
| `columns`         | `columns`     | 递归归一为 `ColumnConfig`           |
| `data`            | `data`        | 原样透传                            |
| `sheetName`       | `name`        | 默认 `"Sheet1"`                     |
| `freezeRows`      | `freezeRows`  | 有定义时透传                        |
| `autoFilter`      | `autoFilter`  | 有定义时透传                        |
| `merges`          | `merges`      | 有定义时透传                        |
| `dataStyle`       | `dataStyle`   | 有定义时透传                        |
| `indexColumn`     | `indexColumn` | 有定义时透传，由 `exportExcel` 展开 |

只有这五个透传字段是「有定义才带上」（`undefined` 仍是 `undefined`，sheet 保持引擎默认）。需要超出这个形状的能力时——多 sheet、表级 `headerStyle`、手写 `ColumnConfig` 树——请用 `tableToSheet` 转换后自己驱动 `exportExcel`（或 [`WorkbookBuilder`](/zh/packages/excel-exporter/api/01-export-excel#其他导出符号)）。

### 完整示例

```ts
import { exportTable, StylePresets } from "@marcusok/excel-exporter";

const result = await exportTable({
  filename: "quarterly-sales",
  sheetName: "Sales",
  freezeRows: 1,
  autoFilter: true,
  indexColumn: { label: "No.", width: 6 },
  dataStyle: StylePresets.dataRow,
  columns: [
    // Element Plus 命名
    { prop: "orderId", label: "Order ID", width: 18 },
    // Ant Design 命名
    { dataIndex: "date", title: "Date", width: 12, format: { type: "date" } },
    // 分组列：多行表头由库自动生成
    {
      label: "Amount",
      headerStyle: StylePresets.header,
      children: [
        { prop: "net", label: "Net", width: 14, style: StylePresets.currency },
        { prop: "tax", label: "Tax", width: 12, style: StylePresets.currency },
      ],
    },
    {
      prop: "region",
      label: "Region",
      width: 12,
      format: {
        type: "enum",
        map: { apac: "APAC", emea: "EMEA", na: "North America" },
        fallback: "Other",
      },
    },
  ],
  data: [
    {
      orderId: "ORD-000001",
      date: "2026-07-01",
      net: 1099.99,
      tax: 199.99,
      region: "apac",
    },
    {
      orderId: "ORD-000002",
      date: "2026-07-02",
      net: 399,
      tax: 72,
      region: "emea",
    },
  ],
});

if (!result.success) console.error(result.error);
```

## exportEcharts

```ts
exportEcharts(options: EChartsExportOptions): Promise<ExportResult>
```

由 ECharts option 的一个小型、明确的结构子集导出一张 sheet。适配层不依赖 ECharts 运行时及其类型系统，因此可以直接把已经构造好的 option 对象传进来——但只认文档化的子集，子集之外一律明确拒绝，不做猜测。

### EChartsExportOptions

`EChartsExportOptions` 在 `EChartsSheetInput` 之上追加了导出级字段：

| 字段             | 类型                                               | 必填 | 说明                                        |
| ---------------- | -------------------------------------------------- | ---- | ------------------------------------------- |
| `option`         | `EChartsOptionInput`                               | 是   | 待转换的（ECharts option 子集）             |
| `filename`       | `string`                                           | 是   | 下载文件名，不以 `.xlsx` 结尾时末尾自动追加 |
| `sheetName`      | `string`                                           | —    | 默认 `"图表数据"`                           |
| `layout`         | `"wide" \| "long"`                                 | —    | 默认 `"wide"`，仅对类目布局有意义           |
| `categoryHeader` | `string`                                           | —    | 默认 `"类目"`                               |
| `seriesHeader`   | `string`                                           | —    | 默认 `"系列"`                               |
| `nameHeader`     | `string`                                           | —    | 默认 `"名称"`                               |
| `valueHeader`    | `string`                                           | —    | 默认 `"数值"`                               |
| `mode`           | `"auto" \| "main" \| "worker" \| "stream"`         | —    | 默认 `"auto"`                               |
| `onProgress`     | `(progress: number) => void`                       | —    | 0 → 1，所有路径上首尾 0 与 1 各上报一次     |
| `onPhase`        | `(phase: ExportPhase, durationMs: number) => void` | —    | `init` / `build` / `download` 阶段耗时      |
| `download`       | `boolean`                                          | —    | 默认 `true`；`false` 只返回 Blob            |

### EChartsOptionInput 及相关类型

| 类型                 | 定义                                                                                                                                                       |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `EChartsScalar`      | `number \| string \| null`                                                                                                                                 |
| `EChartsDatum`       | `EChartsScalar \| number[] \| { name?: string; value?: EChartsScalar \| number[] }`                                                                        |
| `EChartsSeriesInput` | `{ name?: string; type?: string; data?: EChartsDatum[] }`                                                                                                  |
| `EChartsXAxisInput`  | `{ type?: string; data?: Array<number \| string> }`                                                                                                        |
| `EChartsOptionInput` | `{ xAxis?: EChartsXAxisInput \| EChartsXAxisInput[]; yAxis?: EChartsXAxisInput \| EChartsXAxisInput[]; series?: EChartsSeriesInput[]; dataset?: unknown }` |
| `EChartsLayout`      | `"wide" \| "long"`                                                                                                                                         |

关于这些形状的说明：

- 系列与轴上的 `type` 字段会被**忽略**——布局由数据决定，而不是由声明的图表类型决定。
- `yAxis` 复用 `EChartsXAxisInput`：结构上轴就是轴。当 `xAxis.data` 缺失或为空时，它作为类目来源被检索，这正是水平条形图能工作的原因。
- `dataset` 之所以被声明，只是为了能明确地拒绝它。
- 系列名缺失、为空串、`null` 或 `undefined` 时按 `系列1`、`系列2`……（从 1 开始）兜底——空串与未命名同等对待，这样报错总能指到一个具体系列，而不是让用户无从定位。

### 支持的数据形态

**类目布局**——两根轴之一带 `.data`，且每个系列的 `data` 是一维、长度相同：

| 形态                         | 示例                                                                             | 结果（默认 `wide`）           |
| ---------------------------- | -------------------------------------------------------------------------------- | ----------------------------- |
| 柱状/折线（类目在 x 轴）     | `xAxis: { data: ["Q1", "Q2"] }`、`series: [{ name: "Revenue", data: [10, 20] }]` | 每类目一行，每系列一列        |
| 水平条形图（类目在 `yAxis`） | `yAxis: { data: [...] }`，无 `xAxis.data`                                        | 同上——`yAxis.data` 即类目来源 |

类目来源的判定是：`xAxis.data` 为数组且至少一项时取它，否则取 `yAxis.data`。空的 `xAxis.data`（图表尚在加载的中间态）视为「未提供」，回退到 `yAxis.data`，而不是把类目静默丢掉。

**item 布局**——两根轴都没有提供类目时启用。具体走哪种 item 布局，由数据逐系列决定：

| 形态              | 系列数据示例                                                     | 产出的列                                                      |
| ----------------- | ---------------------------------------------------------------- | ------------------------------------------------------------- |
| 饼图 / name-value | `[{ name: "Chrome", value: 62 }, { name: "Safari", value: 19 }]` | `系列` / `名称` / `数值`                                      |
| 裸标量            | `[62, 19]`                                                       | `系列` / `名称` / `数值`（名称取 1 起的序号，`"1"`、`"2"`……） |
| 散点坐标对        | `[[10, 20], [30, 40]]`                                           | `系列` / `X` / `Y`                                            |
| 散点含额外维度    | `[[10, 20, 5], [30, 40, 8]]`                                     | `系列` / `X` / `Y`，额外维度丢弃并以一次 `console.warn` 告知  |
| 散点对象写法      | `[{ value: [10, 20] }, { value: [30, 40] }]`                     | `系列` / `X` / `Y`                                            |

item 布局的细节：

- 散点数据与 name/value 数据不能混用——见下面的拒绝表。
- 散点的两种写法（`[x, y]` 与 `{ value: [x, y] }`）可以互换，也可以彼此混用。对象写法上可选的 `name` 会被接受并忽略：散点布局没有名称列。
- 坐标对的定义是「长度 ≥ 2 且全为数字的数组」。多维散点（`[x, y, ...dims]`，额外的维度驱动 `symbolSize` / `visualMap`）只导出前两维，丢弃行为以**每次导出一次**的 `console.warn` 告知，既不静默、也不是每个点刷一条：
  `[excel-exporter] scatter data has dimensions beyond [x, y]; only the X/Y coordinates are exported (extra dims drive symbolSize/visualMap and have no table column).`
- 只有部分系列带坐标、另一部分带 name/value 数据时，导出会被拒绝（见下）；「是否散点」的判定以「确实存在坐标项」为准，因此散点系列旁边带一个 `data` 为空的系列，不会把饼图误判成散点表。

### 明确拒绝的形态

这些形态会被明确报错拒绝，而不是产出一张会误导人的表。所有文案都以 `[excel-exporter]` 开头：

| 形态                                       | 错误文案                                                                                                                                                                                                                                            |
| ------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 存在 `option.dataset`（dataset 模式）      | `ECharts dataset mode is not supported by echartsToSheet. Flatten the data before calling it.`                                                                                                                                                      |
| `option.series` 缺失、为空、或不是对象数组 | `ECharts option has no series to export.` / `ECharts option.series must be an array of series objects (found a non-object element).`                                                                                                                |
| `xAxis` 或 `yAxis` 多于一根                | `multiple x axes are not supported by echartsToSheet (a dual-axis chart has no single-columnar table shape).`（y 轴则是 `y`）                                                                                                                       |
| 某系列长度与类目轴不一致                   | `ECharts series "系列1" must have the same length as the category axis data (xAxis.data / yAxis.data).`                                                                                                                                             |
| 类目轴下出现对象 / 坐标对数据点            | `unsupported ECharts datum for category layout in series 0 (object/array data points are not supported with a category axis). Flatten each datum to a scalar, or drop xAxis.data/yAxis.data to export via the item (name/value or scatter) layout.` |
| 同一份 option 里散点与 name/value 数据混用 | `mixing scatter coordinate data with name/value data is not supported by echartsToSheet.`                                                                                                                                                           |

拒绝双轴是刻意的：两根轴时不存在单一的表格形状，静默取第一根轴只会把另一根轴的系列报成一条令人费解的「长度不匹配」。

注意类目布局对数据点的限制**仍然允许标量**，包括像 `[5]` 这样的单元素数组（只有长度 ≥ 2 且全为数字的数组才算坐标对）。可靠的经验法则是：类目轴下的系列数据必须是扁平标量。

### layout：wide 与 long

`layout` **只**被类目布局读取。item 数据（饼图 / 散点）完全忽略它，因为这些布局本来就是一条数据点一行。

| `layout`         | 形状                                                                       | 列                                    |
| ---------------- | -------------------------------------------------------------------------- | ------------------------------------- |
| `"wide"`（默认） | 每个类目一行，每个系列一列                                                 | `类目` + 每系列一列（以系列名作表头） |
| `"long"`         | 每个「系列-类目」对一行（`{ seriesHeader, categoryHeader, valueHeader }`） | `系列`、`类目`、`数值`                |

系列不多时，wide 更贴近表格阅读者的直觉；数据透视、筛选、以及系列很多的图，long 更好用。

### 表头

默认表头与 sheet 名是中文，需要英文工作簿时可通过 `sheetName` / `seriesHeader` / `categoryHeader` / `nameHeader` / `valueHeader` 覆盖。随之而来有两条约束：

- **long 与 item 布局要求表头文案互不相同**，因为表头字符串本身兼任行键（`{ [seriesHeader]: name, [categoryHeader]: category, [valueHeader]: value }`）。重复会静默覆盖某一列，因此提前拒绝——例如 `duplicate header "系列" in category long layout: header texts double as row keys in long/item layouts, so they must be distinct (rename via the *Header options).`（item 两种布局的上下文串分别是 `name/value layout` 与 `scatter layout`）。
- **wide 布局拒绝与内部系列键冲突的 `categoryHeader`**：wide 的行以 `categoryHeader` 加上 `__series_0`、`__series_1`…… 为键，因此 `categoryHeader` 若恰好等于其中之一，会被系列数据覆盖。该情况以 `categoryHeader "__series_0" collides with the internal series keys (__series_N) in wide layout; choose a different categoryHeader.` 拒绝。

散点布局的坐标表头是字面量 `X` / `Y`，不可配置；该布局只有 `seriesHeader` 生效。

### 完整示例

```ts
import { exportEcharts } from "@marcusok/excel-exporter";

// 你本来就在渲染的那份图表配置——只会读取文档化的子集。
const option = {
  xAxis: { type: "category", data: ["Q1", "Q2", "Q3", "Q4"] },
  series: [
    { name: "Revenue", type: "bar", data: [120, 200, 150, 80] },
    { name: "Cost", type: "bar", data: [90, 140, 120, 60] },
  ],
};

const result = await exportEcharts({
  filename: "quarterly-chart",
  option,
  sheetName: "Chart data",
  layout: "long", // 每个「系列-类目」对一行
  seriesHeader: "Series",
  categoryHeader: "Quarter",
  valueHeader: "Amount",
});

if (!result.success) console.error(result.error);
```

同一份 option 若用默认值（`layout: "wide"`、中文表头），导出的是名为 `图表数据` 的 sheet，列为 `类目` / `Revenue` / `Cost`。

## 底层转换器

这四个函数都从主入口导出。它们是纯函数——不下载、不挂遮罩、背后也不引用 `exportExcel`——因此当你想要在驱动更底层入口之前先检查、断言或二次加工生成的配置时，它们就是合适的接缝。

| 函数                                                  | 输入                   | 输出            | 适用场景                                                                                         |
| ----------------------------------------------------- | ---------------------- | --------------- | ------------------------------------------------------------------------------------------------ |
| `tableToSheet(input: TableSheetInput)`                | `TableSheetInput`      | `SheetConfig`   | 想用表格形态但换个工作簿装配方式：拼多张 sheet，或喂给 `WorkbookBuilder` / `exportAsStream`      |
| `tableExportToOptions(input: TableExportOptions)`     | `TableExportOptions`   | `ExportOptions` | 想在调用 `exportExcel` 前改动生成的 `ExportOptions`（例如再加一张 sheet，或补一个 `overlay` 值） |
| `echartsToSheet(input: EChartsSheetInput)`            | `EChartsSheetInput`    | `SheetConfig`   | 同 `tableToSheet`，对象是图表数据                                                                |
| `echartsExportToOptions(input: EChartsExportOptions)` | `EChartsExportOptions` | `ExportOptions` | 同 `tableExportToOptions`，对象是图表数据                                                        |

自己驱动更底层入口时有两个注意事项：

- `WorkbookBuilder.addSheet()` 与 `exportAsStream()` **不展开** `indexColumn`；转换后的 sheet 若用到它，请传 `applyIndexColumn(sheet)`（同样已导出）——见 [IndexColumnOptions](/zh/packages/excel-exporter/api/02-types#indexcolumnoptions)。
- 转换器搬运的字段与两个封装一致——包括 `overlay` 不在其中这一点，所以你拿到的就是一份纯粹的配置。

```ts
import { tableToSheet, exportExcel } from "@marcusok/excel-exporter";

const sheet = tableToSheet({
  sheetName: "Sales",
  columns: [{ prop: "region", label: "Region" }],
  data: [{ region: "APAC" }],
});

// 检查或调整生成的配置，再和其他 sheet 一起导出。
await exportExcel({
  filename: "combined",
  sheets: [
    sheet,
    { name: "Notes", columns: [{ prop: "body", label: "Note" }], data: [] },
  ],
  overlay: false, // 只有走主入口才够得着这个选项
});
```

## 错误契约

两个封装在转换错误上从不 reject。它们的选项类型承诺的是 `Promise<ExportResult>`，输入里的结构性问题——列没有名字、出现 `dataset`、系列长度不匹配、表头重复——都会以普通的失败结果返回：

```ts
const result = await exportEcharts({ filename: "x", option: { dataset: {} } });

result.success; // false
result.error; // Error: [excel-exporter] ECharts dataset mode is not supported ...
```

这与 `exportExcel` 一致——主入口本就以 `{ success: false, error }` 上报自身失败：转换错误只是发生在进入 `exportExcel` **之前**，于是封装把它们捕获并包装成同一种形态。因此同一份坏输入，无论走便捷封装还是手工拼 `ExportOptions`，失败形态都相同。

还需注意 `success: false` 也可能来自兜底路径——判断时请看 `result.error`（以及 `result.engine` / `result.mode`），不要假定用了某个特定引擎。另见 [ExportResult](/zh/packages/excel-exporter/api/01-export-excel#exportresult)。
