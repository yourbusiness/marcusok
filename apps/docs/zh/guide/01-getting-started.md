# 快速开始

MarcusOK 共发布两个包、分属两个大类——它们之间的关系见 [生态介绍](/zh/guide/)。本文先把最先要决定的事说清楚（该装哪个包），再用其中最常用的那个带你跑通第一次导出。

运行环境要求：Node `>= 22`。示例命令使用 pnpm，用 npm / yarn 安装等价（`pnpm >= 9` 只是本仓库自身的开发环境要求，与消费方无关）。

## 该装哪个包？

| 我要做的事                                    | 安装                       | 从这里开始                                                      |
| --------------------------------------------- | -------------------------- | --------------------------------------------------------------- |
| 把数据导出成可下载的 `.xlsx`                  | `@marcusok/excel-exporter` | 下方的[导出演练](#_2-第一个导出)                                |
| 只读预览用户上传的 `.xlsx` / `.xlsm` / `.csv` | `@marcusok/excel-preview`  | [预览快速开始](/zh/packages/excel-preview/guide/01-quick-start) |

下文演练走 `@marcusok/excel-exporter`；预览包的快速开始见上表链接。两件曾以独立包形式提供的事：进度遮罩经导出包的 `/overlay` 子路径取用（见[进度遮罩指南](/zh/packages/excel-exporter/guide/11-overlay)）；独立的引擎包已不存在——两个业务包各自内置引擎并再导出其加载器。完整对比——依赖图、运行时各自内置什么、版本如何对应——见[包关系与选型](/zh/guide/03-package-relationships)。

## 1. 安装

```bash
pnpm add @marcusok/excel-exporter
```

这就是全部——没有需要额外操心的同 scope 运行时依赖：导出引擎（modern-xlsx JS 胶水 + fflate）与进度遮罩 UI 都在构建期打包进本包 `dist`，WASM 二进制通过本包自己的 `exports` 映射分发。既没有需要额外接线的引擎包，也不会被上游的 engines 声明影响安装。

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
        { prop: "orderId", label: "订单号", width: 18 },
        { prop: "date", label: "日期", width: 12, format: { type: "date" } },
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

无需在 `main.ts` 接线、无需打包器插件：随包发布的资产（本包自己的 `modern-xlsx.wasm` 与 `export.worker.js`）默认自动定位——打包器通过标准 `new URL(资产, import.meta.url)` 模式把它们发射为 hash 资产，Node 直接从磁盘读取 wasm。Node / SSR 环境同样零配置（详见 [Node/SSR](/zh/packages/excel-exporter/guide/09-node-ssr)）。只有自托管 / CDN 托管副本的场景才需要 [`configureWasm`](/zh/packages/excel-exporter/guide/02-installation)。

## 3. 下一步

- 了解 [自动模式路由](/zh/packages/excel-exporter/guide/03-auto-mode) 是如何选择 main / worker / stream 的
- 在 [在线演示](/zh/play) 里直接体验不同数据量与模式，同一页还有预览包与进度遮罩的演示
- 从上表挑另一个包开始，或看[包关系与选型](/zh/guide/03-package-relationships)做完整对比
