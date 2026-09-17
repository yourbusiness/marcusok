# 案例：销售月报导出

月度销售报表是后台系统最高频的导出场景。本案例用 mock 生成 1 万行销售明细，演示：日期/金额/枚举格式化、货币样式、冻结表头、自动筛选与进度回调。

## Mock 数据预览

<MockPreview dataset="sales" :rows="5" />

## 实现代码

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

// rows：业务侧的销售明细数据，字段与下方 columns 一一对应
//（由你的业务代码提供，此处省略取数过程；示例场景为 1 万行，
//  在线演示用的是同一销售场景的 mock 数据，比示例多一列 unitPrice）

const result = await exportExcel({
  filename: "销售月报-2026-07",
  sheets: [
    {
      name: "销售明细",
      freezeRows: 1,
      autoFilter: true,
      columns: [
        { prop: "orderId", label: "订单号", width: 18 },
        {
          prop: "date",
          label: "日期",
          width: 12,
          format: { type: "date" },
        },
        { prop: "region", label: "区域", width: 10 },
        { prop: "product", label: "商品", width: 18 },
        { prop: "channel", label: "渠道", width: 10 },
        { prop: "quantity", label: "数量", width: 8 },
        {
          prop: "amount",
          label: "金额",
          width: 14,
          style: StylePresets.currency,
        },
        {
          prop: "status",
          label: "状态",
          width: 10,
          format: {
            type: "enum",
            map: { paid: "已支付", pending: "待支付", refunded: "已退款" },
            fallback: "未知",
          },
        },
      ],
      data: rows,
    },
  ],
  onProgress: (p) => setProgress(p), // 展示进度条
  onPhase: (phase, ms) => trackPhase(phase, ms), // 埋点
});
```

## 要点

- 1 万行低于 20,000 行阈值，浏览器下走 `main` 路径（主线程 Workbook，样式完整保留）；数据量 ≥ 20,000 行时自动切换到 `worker + Workbook`，主线程不阻塞；
- `onProgress` 在 `main` 路径只回调首尾（0 与 1），1 万行下进度条会直接跳满；需要分段进度必须走 stream 路径（≥ 50,000 行，每 1000 行上报一次；worker + Workbook 路径同样只有首尾两次回调）；
- 金额列用 `StylePresets.currency`（千分位 + 两位小数，右对齐）；
- 状态列用 `enum` 把内部码映射为中文，兜底 `"未知"`；
- `freezeRows + autoFilter` 让管理层在 Excel 里直接筛选。

可以到 [在线演示](/zh/play) 选择 sales 数据集直接导出体验。
