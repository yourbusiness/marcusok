# API：StylePresets 样式预设

## 预设一览

| 名称       | numFormat          | 其他样式                                   | 适用               |
| ---------- | ------------------ | ------------------------------------------ | ------------------ |
| `header`   | —                  | 加粗、12 号字、深蓝底 `1F4E79`、白字、居中 | 表头（手动应用时） |
| `currency` | `#,##0.00`         | 右对齐                                     | 金额               |
| `percent`  | `0.00%`            | 右对齐                                     | 占比、增长率       |
| `date`     | `yyyy-MM-dd`       | 居中                                       | 日期列             |
| `datetime` | `yyyy-MM-dd HH:mm` | 居中                                       | 日期时间列         |
| `dataRow`  | —                  | 左对齐、垂直居中、底部细线 `D0D0D0`        | 数据行             |
| `bordered` | —                  | 四边细线 `D0D0D0`                          | 整表边框           |
| `danger`   | —                  | 红色加粗 `C00000`、居中                    | 风险/异常值        |

## 完整定义

预设是普通的常量对象——即 `src/style-presets.ts` 中随包发布的真实值：

```ts
export const StylePresets = {
  header: {
    font: { bold: true, size: 12, color: "FFFFFF" },
    fill: { pattern: "solid", fgColor: "1F4E79" },
    alignment: { horizontal: "center", vertical: "center" },
  },
  currency: { numFormat: "#,##0.00", alignment: { horizontal: "right" } },
  percent: { numFormat: "0.00%", alignment: { horizontal: "right" } },
  date: { numFormat: "yyyy-MM-dd", alignment: { horizontal: "center" } },
  datetime: {
    numFormat: "yyyy-MM-dd HH:mm",
    alignment: { horizontal: "center" },
  },
  dataRow: {
    alignment: { horizontal: "left", vertical: "center" },
    border: { bottom: { style: "thin", color: "D0D0D0" } },
  },
  bordered: {
    border: {
      top: { style: "thin", color: "D0D0D0" },
      bottom: { style: "thin", color: "D0D0D0" },
      left: { style: "thin", color: "D0D0D0" },
      right: { style: "thin", color: "D0D0D0" },
    },
  },
  danger: {
    font: { color: "C00000", bold: true },
    alignment: { horizontal: "center" },
  },
} as const;
```

## 用法

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

columns: [
  { prop: "amount", label: "金额", width: 14, style: StylePresets.currency },
  { prop: "rate", label: "增长率", width: 12, style: StylePresets.percent },
  { prop: "date", label: "日期", width: 12, style: StylePresets.date },
  { prop: "flag", label: "状态", width: 10, style: StylePresets.danger },
];
```

## 样式挂载位置

| 目标                     | 字段                                        | 合并语义                        |
| ------------------------ | ------------------------------------------- | ------------------------------- |
| 某列的数据单元格         | `ColumnConfig.style`                        | 与表级 `dataStyle` 逐字段深合并 |
| 某列的表头（含分组表头） | `ColumnConfig.headerStyle`                  | 整体替换表级 `headerStyle`      |
| 全部数据单元格（基底）   | `SheetConfig.dataStyle`                     | 被列级 `style` 逐字段覆盖       |
| 全部表头格（默认）       | `SheetConfig.headerStyle`                   | 被列级 `headerStyle` 整体替换   |
| 注入的序号列             | `IndexColumnOptions.style` / `.headerStyle` | 与普通列规则一致                |

## 派生变体

预设不是工厂函数、也没有冻结——用对象展开派生变体：

```ts
// 无小数的金额
style: { ...StylePresets.currency, numFormat: "#,##0" }

// 只改嵌套对象的一个字段：嵌套对象也要展开，
// 否则同级字段会全部丢失（展开是浅合并）
style: { ...StylePresets.header, font: { ...StylePresets.header.font, size: 14 } }
```

展开是调用方自己的合并，发生在库读取配置之前——与引擎在 `dataStyle` 与列级 `style` 之间的字段级深合并是两回事。现成配方（报表模板、表头覆盖、`numFormat` 调整）见[样式指南](/zh/packages/excel-exporter/guide/05-styles)。

## 类型

```ts
import type { StylePresetName } from "@marcusok/excel-exporter";

const name: StylePresetName = "currency"; // "header" | "currency" | "percent" | "date" | "datetime" | "dataRow" | "bordered" | "danger"
```

`StylePresetName` 用于参数化预设选择：

```ts
const columnPresets: Record<string, StylePresetName> = {
  amount: "currency",
  rate: "percent",
  orderDate: "date",
};

const columns = Object.entries(columnPresets).map(([prop, name]) => ({
  prop,
  label: prop,
  style: StylePresets[name],
}));
```

> 注意：`style` 应用于数据单元格，不作用于表头。需要表头样式时，可直接用 `headerStyle` 字段（工作表级 `SheetConfig.headerStyle` 设默认，列级 `ColumnConfig.headerStyle` 覆盖），例如 `headerStyle: StylePresets.header`。
