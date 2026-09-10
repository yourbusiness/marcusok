# 快速开始

本文带你从零跑通第一个 Excel 导出。运行环境要求：Node `>= 22`。示例命令使用 pnpm，用 npm / yarn 安装等价（`pnpm >= 9` 只是本仓库自身的开发环境要求，与消费方无关）。

## 1. 安装

```bash
pnpm add @marcusok/excel-exporter
```

这就是全部——一个包、**零运行时依赖**：导出引擎（modern-xlsx JS 胶水 + fflate）已在构建期打包进来，WASM 二进制通过本包自己的 `exports` 暴露，既没有需要额外安装的引擎包，也不会被上游的 engines 声明影响安装。

## 2. 第一个导出

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "销售明细-2026",
  sheets: [
    {
      name: "销售明细",
      freezeRows: 1,
      autoFilter: true,
      columns: [
        { key: "orderId", header: "订单号", width: 18 },
        { key: "date", header: "日期", width: 12, format: { type: "date" } },
        {
          key: "amount",
          header: "金额",
          width: 14,
          style: StylePresets.currency,
        },
        {
          key: "status",
          header: "状态",
          width: 10,
          format: {
            type: "enum",
            map: { paid: "已支付", pending: "待支付" },
            fallback: "未知",
          },
        },
      ],
      data: [
        {
          orderId: "ORD-000001",
          date: "2026-07-01",
          amount: 1299.99,
          status: "paid",
        },
      ],
    },
  ],
});
```

浏览器中运行会自动触发下载，文件名缺省 `.xlsx` 后缀时自动补全。`download: false` 时只返回 Blob，便于自托管上传等场景。

无需在 `main.ts` 接线、无需打包器插件：随包发布的两份资产（`modern-xlsx.wasm`、`export.worker.js`）默认自动定位——打包器通过标准 `new URL(资产, import.meta.url)` 模式把它们发射为 hash 资产，Node 直接从磁盘读取 wasm。Node / SSR 环境同样零配置（详见 [Node/SSR](/zh/packages/excel-exporter/guide/09-node-ssr)）。只有自托管 / CDN 托管副本的场景才需要 [`configureWasm`](/zh/packages/excel-exporter/guide/02-installation)。

## 3. 下一步

- 了解 [自动模式路由](/zh/packages/excel-exporter/guide/03-auto-mode) 是如何选择 main / worker / stream 的
- 在 [在线演示](/zh/play) 里直接体验不同数据量与模式
