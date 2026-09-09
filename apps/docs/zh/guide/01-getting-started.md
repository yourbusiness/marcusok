# 快速开始

本文带你从零跑通第一个 Excel 导出。运行环境要求：Node `>= 22`。示例命令使用 pnpm，用 npm / yarn 安装等价（`pnpm >= 9` 只是本仓库自身的开发环境要求，与消费方无关）。

## 1. 安装

```bash
pnpm add @marcusok/excel-exporter
```

> **Node 版本说明**：引擎依赖 `modern-xlsx` 会自动安装，它声明了 `engines.node >= 24`，但其 WASM 核心面向浏览器，Node 22 实际可正常使用——本包在 Node 22 上开发并通过 CI。若包管理器在 Node 22 下因 engines 校验拒绝安装（如 pnpm 开启 `engine-strict` 时报错），可在项目 `.npmrc` 中加一行 `engine-strict=false`，或升级到 Node ≥ 24。

`modern-xlsx` 是直接 `dependency`：本包锁定经过测试的引擎版本，并把 `modern-xlsx.wasm` 在构建时转发进自己的 `dist/` 通过 `exports` 暴露（`@marcusok/excel-exporter/dist/modern-xlsx.wasm`），保证二进制与 JS 胶水版本一致，也让打包器可以直接 `?url` 导入（modern-xlsx 自身的 `exports` 未暴露 wasm 子路径）。

`xlsx`（SheetJS）为**可选** peerDependency，仅在需要降级兜底时安装；不安装时兜底会从官方 CDN 加载。

## 2. 配置浏览器静态资源（仅浏览器需要）

一份资源在绝大多数路径下都需要：

- `modern-xlsx.wasm`（约 1.9MB，样式引擎核心；唯一例外是显式 `mode: "stream"` 的纯 JS 流式路径）

另一份 `export.worker.js` 只在导出真正进入 Worker 时才需要——auto 模式 ≥ 20,000 行，或显式 `mode: "worker"` / `mode: "stream"`。两份文件都在本包 `dist/` 里，全部通过 `@marcusok/excel-exporter` 的导入解析。

### Vite（推荐）

用 `?url` 后缀导入资源——Vite 开发态直接服务、构建时自动 hash 拷贝进 `dist/assets/`，无需拷贝插件、无需 `public/` 目录：

```ts
// main.ts
import { configureWasm } from "@marcusok/excel-exporter";
// 由本包转发暴露（modern-xlsx 自身的 exports 未包含 wasm 子路径）
import wasmUrl from "@marcusok/excel-exporter/dist/modern-xlsx.wasm?url";
// 可选——导出行数始终 < 2 万时，删掉这行即可
import workerUrl from "@marcusok/excel-exporter/dist/export.worker.js?url";

configureWasm({ wasmUrl, workerUrl });
```

### 其他构建方案（拷贝文件）

打包器不支持资产 URL 导入时，从本包 `dist/` 把两份文件拷贝到静态目录。在 `buildStart` 中用 `require.resolve` 反推真实路径，避免硬编码 `node_modules`（pnpm 符号链接不兼容）：

```ts
// vite.config.ts / 构建脚本
import { defineConfig } from "vite";
import { createRequire } from "node:module";
import { copyFileSync, mkdirSync, statSync } from "node:fs";
import { dirname } from "node:path";

const require = createRequire(import.meta.url);
// 唯一来源：两份资产都在 @marcusok/excel-exporter/dist
const pkgDist = dirname(require.resolve("@marcusok/excel-exporter"));

export default defineConfig({
  plugins: [
    {
      name: "copy-excel-exporter-assets",
      buildStart() {
        mkdirSync("public/assets", { recursive: true });
        for (const file of ["modern-xlsx.wasm", "export.worker.js"]) {
          const src = `${pkgDist}/${file}`;
          if (!statSync(src, { throwIfNoEntry: false }))
            throw new Error(`${file} not found. Looked at: ${src}`);
          copyFileSync(src, `public/assets/${file}`);
        }
      },
    },
  ],
});
```

```ts
// main.ts
import { configureWasm } from "@marcusok/excel-exporter";

configureWasm({
  wasmUrl: "/assets/modern-xlsx.wasm",
  workerUrl: "/assets/export.worker.js",
});
```

Node / SSR 环境**无需**部署浏览器静态资源，也**无需**任何初始化样板：未配置 `wasmUrl` 时，引擎会在首次使用时自动从 `node_modules` 定位并同步初始化 WASM（详见 [Node/SSR](/zh/packages/excel-exporter/guide/09-node-ssr)）。

## 3. 第一个导出

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

## 4. 下一步

- 了解 [自动模式路由](/zh/packages/excel-exporter/guide/03-auto-mode) 是如何选择 main / worker / stream 的
- 在 [在线演示](/zh/play) 里直接体验不同数据量与模式
