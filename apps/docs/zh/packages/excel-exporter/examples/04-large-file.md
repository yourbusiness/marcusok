# 案例：10 万行大文件导出

大数据量导出是 Fast stream 路径的目标场景：10 万行在浏览器 Worker 中执行，主线程仅做一次结构化克隆，Fast stream 约 0.8s 完成。

## Mock 数据预览

<MockPreview dataset="sales" :rows="3" />

## 实现代码

```ts
import { exportExcel } from "@marcusok/excel-exporter";

// rows：业务侧的销售明细数据，字段与下方 columns 一一对应
//（由你的业务代码提供，此处省略取数过程；示例场景为 10 万行，
//  在线演示用的是同一销售场景的 mock 数据，列数为完整列集）

const result = await exportExcel({
  filename: "大文件导出-10w",
  sheets: [
    {
      name: "销售明细",
      columns: [
        { prop: "orderId", label: "订单号", width: 18 },
        { prop: "date", label: "日期", width: 12 },
        { prop: "amount", label: "金额", width: 14 },
        { prop: "status", label: "状态", width: 10 },
      ],
      data: rows,
    },
  ],
  mode: "auto", // 达到行/单元格 stream 门槛后自动走 worker + stream
  onProgress: (p) => setProgress(p),
});

console.log(result); // engine: "fast-xlsx", mode: "stream", rowCount: 100000
```

## 要点

- 10 万行下 `auto` 选择 `worker + Fast stream`，实测约 0.8s（Workbook 路径同数据量 Chrome 实测 17.5s）；
- stream 路径**支持多行表头、合并与数字格式**（date/number FormatSpec 经最小 `styles.xml`），但**不支持单元格/布局样式、列宽/冻结/筛选**，会打印 console 警告，属预期行为；
- `onProgress` 每 1000 行上报一次，适合展示进度；
- `decimals`/`thousands` 在两条路径上都是纯显示 numFmt，跨阈值存储值保持一致。

到 [在线演示](/zh/play) 选择 100,000 行，切换 `auto` 与 `main` 对比体验差异。
