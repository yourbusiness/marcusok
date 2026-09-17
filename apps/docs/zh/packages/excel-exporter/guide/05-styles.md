# 样式

列级 `style`（`CellStyle`）应用于该列的所有**数据单元格**——表头样式单独配置：列级 `headerStyle` 或表级 `headerStyle`；表级 `dataStyle` 则为全部数据单元格提供基底样式（见下文）。内置 8 种预设，也支持完全自定义——包括用对象展开（spread）从预设派生自己的变体。

## 内置预设 StylePresets

| 预设                    | 视觉效果                                                                                                          | 说明                                                     |
| ----------------------- | ----------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------- |
| `StylePresets.header`   | <span style="display:inline-block;width:12px;height:12px;background:#1F4E79;border-radius:2px"></span> 深蓝底白字 | 加粗、12 号字、深蓝底 `1F4E79`、白字、居中               |
| `StylePresets.currency` | `#,##0.00`                                                                                                        | 千分位 + 两位小数，右对齐                                |
| `StylePresets.percent`  | `0.00%`                                                                                                           | 百分比格式，右对齐                                       |
| `StylePresets.date`     | `yyyy-MM-dd`                                                                                                      | 日期格式，居中                                           |
| `StylePresets.datetime` | `yyyy-MM-dd HH:mm`                                                                                                | 日期时间格式，居中                                       |
| `StylePresets.dataRow`  | 左对齐 + 底部细线                                                                                                 | 左对齐、垂直居中，底部浅灰细线 `D0D0D0`                  |
| `StylePresets.bordered` | 四边细线框                                                                                                        | 四边浅灰细线 `D0D0D0`，适合配表级 `dataStyle` 做整表边框 |
| `StylePresets.danger`   | <span style="display:inline-block;width:12px;height:12px;background:#C00000;border-radius:2px"></span> 红色加粗   | 红色加粗文字 `C00000`，居中                              |

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "styled",
  sheets: [
    {
      name: "Sheet1",
      columns: [
        { prop: "name", label: "名称", width: 16, style: StylePresets.dataRow },
        {
          prop: "amount",
          label: "金额",
          width: 14,
          style: StylePresets.currency,
        },
        { prop: "date", label: "日期", width: 12, style: StylePresets.date },
        {
          prop: "status",
          label: "状态",
          width: 10,
          style: StylePresets.danger,
        },
      ],
      data: [
        {
          name: "机械键盘",
          amount: 1299.99,
          date: "2026-07-01",
          status: "缺货",
        },
      ],
    },
  ],
});
```

## 组合使用

预设为组合而设计：一个表级 `headerStyle`、一个表级 `dataStyle`，加上少量列级 `style`，就能产出完整的报表观感：

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "monthly-report",
  sheets: [
    {
      name: "Orders",
      headerStyle: StylePresets.header, // 整表深蓝表头
      dataStyle: StylePresets.bordered, // 每个数据单元格的细边框
      indexColumn: { label: "No.", width: 6 }, // 最左侧序号列
      freezeRows: 1,
      autoFilter: true,
      columns: [
        { prop: "orderDate", label: "Date", style: StylePresets.date },
        { prop: "amount", label: "Amount", style: StylePresets.currency },
        { prop: "rate", label: "Growth", style: StylePresets.percent },
        { prop: "status", label: "Status", style: StylePresets.danger },
      ],
      data,
    },
  ],
});
```

每个视觉元素的来源：

| 视觉元素                                   | 来源字段           |
| ------------------------------------------ | ------------------ |
| 整表深蓝表头                               | 表级 `headerStyle` |
| 每个数据单元格的细边框                     | 表级 `dataStyle`   |
| 最左侧 `No.` 序号列                        | 表级 `indexColumn` |
| `yyyy-MM-dd` 日期、`#,##0.00` 金额、百分比 | 列级 `style`       |

列级样式逐字段合并**覆盖** `dataStyle`（见下文「表级 dataStyle」）：`StylePresets.currency` 只声明 `numFormat` + `alignment`，因此它的单元格既保留 `dataStyle` 的整表边框，又获得列级数字格式。

## 预设定制（spread 派生）

预设是普通的常量对象（`as const`，经 `satisfies CellStyle` 收窄类型）——不是工厂函数，也没有冻结。与其从零重写 `CellStyle`，不如用对象展开从预设派生变体：

```ts
// 千分位但不要小数
style: { ...StylePresets.currency, numFormat: "#,##0" }

// 金额加粗（currency 未设置 font，不会丢任何东西）
style: { ...StylePresets.currency, font: { bold: true } }

// 同一款表头预设换个品牌色
headerStyle: { ...StylePresets.header, fill: { pattern: "solid", fgColor: "2E7D32" } }

// 序号列：水平居中，同时保留 dataRow 的垂直居中
// （嵌套 alignment 按下方浅合并规则展开）
indexColumn: {
  style: {
    ...StylePresets.dataRow,
    alignment: { ...StylePresets.dataRow.alignment, horizontal: "center" },
  },
}
```

::: warning 对象展开是浅合并
展开会**整体替换**顶层字段。嵌套对象（`font`、`fill`、`alignment`、`border`）必须自行展开，否则同级字段会全部丢失：

```ts
// ✗ 加粗 / 字号 / 颜色全丢——font 被整体替换
{ ...StylePresets.header, font: { size: 14 } }

// ✓ 只改字号，font 其余字段保留
{ ...StylePresets.header, font: { ...StylePresets.header.font, size: 14 } }
```

这是调用方自己的合并，发生在库介入之前——与引擎在 `dataStyle` 与列级 `style` 之间做的字段级**深合并**（见下文）是两回事。
:::

数值类预设常用的 `numFormat` 变体：

| 格式码               | 渲染效果           |
| -------------------- | ------------------ |
| `#,##0`              | `12,999`（无小数） |
| `0%`                 | `42%`（无小数）    |
| `"¥"#,##0.00`        | `¥12,999.99`       |
| `yyyy"年"M"月"d"日"` | `2026年7月1日`     |

格式码中的字面量文本用双引号包裹；其余遵循 Excel 格式码语法。

## 自定义 CellStyle

```ts
import type { CellStyle } from "@marcusok/excel-exporter";

const highlight: CellStyle = {
  font: { bold: true, size: 11, color: "1F4E79" }, // 6 位 RGB hex
  fill: { pattern: "solid", fgColor: "DDEBF7" },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  border: {
    bottom: { style: "medium", color: "1F4E79" },
    right: { style: "thin", color: "D0D0D0" },
  },
  numFormat: "#,##0.00",
};
```

字段约定：

| 字段        | 说明                                                                                                    |
| ----------- | ------------------------------------------------------------------------------------------------------- |
| `font`      | `bold` / `italic` / `size` / `color`（6 位 hex，如 `"FF0000"`）/ `name`                                 |
| `fill`      | `pattern: "solid" \| "none"`、`fgColor`、`bgColor`（6 位 hex）                                          |
| `alignment` | `horizontal`（left/center/right）、`vertical`（top/center/bottom）、`wrapText`、`textRotation`（0–180） |
| `border`    | 四边 `{ style, color }`，`style` 取值见 modern-xlsx 的 `BorderStyle`                                    |
| `numFormat` | Excel 数字格式码，如 `"#,##0.00"`、`"yyyy-mm-dd"`、`"0.00%"`                                            |

> 颜色统一使用 6 位 RGB hex（不带 `#`），与 modern-xlsx 的类型约定一致。

## 表级 dataStyle

`SheetConfig.dataStyle` 是**全部数据单元格**的基底样式——一个字段覆盖整表，无需逐列重复：

```ts
await exportExcel({
  filename: "bordered",
  sheets: [
    {
      name: "Sheet1",
      headerStyle: StylePresets.header, // 表头
      dataStyle: StylePresets.bordered, // 全部数据单元格，一个字段搞定
      columns: [
        { prop: "name", label: "名称" },
        // 列级 style 逐字段深合并覆盖：整表边框保留，
        // 数字格式 / 对齐来自列级配置
        { prop: "amount", label: "金额", style: StylePresets.currency },
      ],
      data,
    },
  ],
});
```

合并是**字段级**的，不是整体替换：`dataStyle` 提供基底，列级 `style` 只覆盖它声明了的字段——整表边框不会被一个只设置了 `numFormat` 的列冲掉，反之亦然。这与 `headerStyle` 的整体替换语义是刻意不同的。与其他样式一样，`dataStyle` 在流式路径（≥ 50,000 行 / 降级导出）会被丢弃并告警。

## 表头样式：表级默认、列级覆盖

表头同样有两个层级，但合并语义相反——**整体替换**，不做字段级合并：

```ts
sheets: [
  {
    name: "Sheet1",
    headerStyle: StylePresets.header, // 所有表头格的默认样式
    columns: [
      { prop: "name", label: "名称" },
      // 列级 headerStyle 对该列表头格整体替换表级默认——
      // 不是与之合并。
      {
        prop: "amount",
        label: "金额",
        headerStyle: {
          fill: { pattern: "solid", fgColor: "DDEBF7" },
          font: { bold: true, color: "1F4E79" },
        },
      },
    ],
    data,
  },
],
```

- 分组列（带 `children`）同样接受 `headerStyle`——作用于该分组横向合并后的表头格。
- 想在默认基础上派生而不是重写时，展开它：`headerStyle: { ...StylePresets.header, fill: { pattern: "solid", fgColor: "DDEBF7" } }`——浅合并注意事项见上文「预设定制」。

## 序号列

`SheetConfig.indexColumn` 在最左侧注入序号列——无需预处理数据，也不用多写一个列配置：

```ts
sheets: [
  {
    name: "Sheet1",
    indexColumn: true, // 或 { label: "序号", width: 6, start: 1, style, headerStyle }
    columns: [
      { prop: "name", label: "名称" },
      { prop: "amount", label: "金额" },
    ],
    data, // 原样传入——序号由行号生成
  },
];
```

- **所有导出路径**（workbook / worker / stream）均支持：它是结构而非样式，连无样式流式路径也保留序号。
- 已有 `merges` 自动右移一列，仍指向原目标。
- 序号值从不读取 `data`；用户列声明保留 prop `__index__` 会被明确报错拒绝。
- 样式规则与普通列一致：表头走 `indexColumn.headerStyle`（优先于表级 `headerStyle`），数据单元格走 `indexColumn.style`（与表级 `dataStyle` 合并）。

完整选项见 [API 参考](../api/02-types)。

## 样式的生效范围

- **样式仅在 Workbook 路径渲染**：auto 模式 50,000 行以下（主线程，或浏览器 20,000–49,999 行的 Worker + Workbook）。流式路径（≥ 50,000 行或降级导出）会剥离全部样式——`style` / `headerStyle` / `dataStyle` / `width` / `freezeRows`——并输出 console 告警。大文件导出不要依赖样式。
- **颜色是 6 位 RGB hex、不带 `#`**（`"1F4E79"`），与 modern-xlsx 的类型约定一致。
- **复用预设零成本**：引擎对结构相同的样式去重、共享同一样式索引，N 个列引用 `StylePresets.currency` 不会膨胀文件——无需手工共享对象。

## 与 FormatSpec 的关系

`style.numFormat` 与列级 `format`（FormatSpec）是两套机制：`style` 控制单元格外观，`format` 负责把业务值转换/格式化为可写入的值。对 `date` / `number` 等 FormatSpec，Workbook 路径会自动注入对应的 `numFormat`，通常不需要手动设置。详见 [值格式化](/zh/packages/excel-exporter/guide/04-formatting)。
