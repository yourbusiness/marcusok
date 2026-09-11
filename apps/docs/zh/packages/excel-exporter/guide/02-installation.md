# 安装与配置

## 环境要求

- Node `>= 22`（包的 `engines` 要求）；包管理器不限，本文示例命令用 pnpm，npm / yarn 等价
- 浏览器需要支持 WebAssembly（现代浏览器均支持）

## 安装

```bash
pnpm add @marcusok/excel-exporter
```

这就是全部。本包**零运行时依赖**：导出引擎（modern-xlsx JS 胶水 + fflate）已在构建期打包进来，WASM 二进制通过本包自己的 `exports` 暴露——没有需要额外安装的引擎包、没有可选兜底包，也不需要在 `main.ts` 或打包器配置里加任何东西。

## 浏览器：资源自动定位

两份文件随包发布、默认自动定位，直接 `import { exportExcel } from "@marcusok/excel-exporter"` 即可使用：

| 资源               | 何时需要                                                                                                                                           |
| ------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| `modern-xlsx.wasm` | 约 1.9MB，样式引擎核心；带样式路径（main / worker + Workbook）需要——Fast stream 路径（显式 `mode: "stream"`，或 auto/Node ≥ 50,000 行）不依赖 WASM |
| `export.worker.js` | 自包含 worker（单文件 ESM，零自身 import）；仅进入 Worker 的路径需要（auto ≥ 20,000 行，以及显式 `mode: "worker"` / `mode: "stream"`）             |

定位顺序：

1. **打包器**——两个 URL 默认为相对包入口的 `new URL(<文件>, import.meta.url)`。Vite 在开发态（依赖预构建）会改写该表达式、构建时发射 hash 资产；webpack 5 文档支持同样的资产模式。无需插件、无需 `?url` 导入、无需拷贝。
2. **Node**——直接从安装目录旁的磁盘读取二进制并同步初始化（见 [Node / SSR](#node-ssr)）。

### 可选：`configureWasm`

默认定位无法生效的场景（CDN 自托管、Service Worker 环境、不支持资产 URL 的打包器）或需要调加载超时的场景，用 `configureWasm` 覆盖：

```ts
import { configureWasm } from "@marcusok/excel-exporter";

configureWasm({
  // 两项都可选，只覆盖需要的
  wasmUrl: "https://cdn.example.com/modern-xlsx.wasm",
  workerUrl: "https://cdn.example.com/export.worker.js",
});
```

支持资产导入的打包器也可以显式接线（完全受支持，2.0 之前的推荐接法）：

```ts
import wasmUrl from "@marcusok/excel-exporter/dist/modern-xlsx.wasm?url";
import workerUrl from "@marcusok/excel-exporter/dist/export.worker.js?url";
configureWasm({ wasmUrl, workerUrl });
```

| 参数              | 类型            | 默认值             | 说明                                             |
| ----------------- | --------------- | ------------------ | ------------------------------------------------ |
| `wasmUrl`         | `string \| URL` | 随包发布的 `.wasm` | 覆盖为自托管 / CDN 副本                          |
| `workerUrl`       | `string \| URL` | 随包发布的 worker  | 覆盖为自托管 / CDN 副本                          |
| `timeoutMs`       | `number`        | `10_000`           | 单次加载超时                                     |
| `maxRetries`      | `number`        | `3`                | 最大尝试次数；失败后按 300ms、600ms 指数退避等待 |
| `workerTimeoutMs` | `number`        | `120_000`          | Worker 导出超时；超时导出会终止共享 worker       |

`configureWasm` 是合并语义：仅当 `wasmUrl` 变化时才重置已加载（或加载中）的 WASM 实例，只改超时/重试不会造成重复初始化；若此前加载失败（error 态），任意 `configureWasm` 调用都会清除错误态，下次导出按新配置重试。

## Node / SSR

Node 环境无需部署浏览器静态资源，也**无需任何初始化样板**：未配置 `wasmUrl` 时，引擎会在首次使用时自动定位安装目录旁的 `dist/modern-xlsx.wasm`（pnpm 符号链接安全）并同步初始化。`auto` 模式下 Node 不会走 Worker，而是主线程执行；≥ 5 万行自动切换流式路径（流式路径不依赖 WASM）。初始化时机控制与打包器注意事项见 [Node/SSR](/zh/packages/excel-exporter/guide/09-node-ssr)。
