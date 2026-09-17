# 样式

列级 `style`（`CellStyle`）应用于该列的所有**数据单元格**——表头样式单独配置：列级 `headerStyle` 或表级 `headerStyle`；表级 `dataStyle` 则为全部数据单元格提供基底样式（见下文）。内置 8 种预设，也支持完全自定义。

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

## 与 FormatSpec 的关系

`style.numFormat` 与列级 `format`（FormatSpec）是两套机制：`style` 控制单元格外观，`format` 负责把业务值转换/格式化为可写入的值。对 `date` / `number` 等 FormatSpec，Workbook 路径会自动注入对应的 `numFormat`，通常不需要手动设置。详见 [值格式化](/zh/packages/excel-exporter/guide/04-formatting)。
