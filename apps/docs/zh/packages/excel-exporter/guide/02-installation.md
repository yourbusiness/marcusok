# 安装与配置

## 环境要求

- Node `>= 22`（包的 `engines` 要求）；包管理器不限，本文示例命令用 pnpm，npm / yarn 等价。依赖 `modern-xlsx` 额外声明了 `engines.node >= 24`——Node 22 实际可用，但开启 engines 校验的包管理器会拒绝安装（见下方说明）
- 浏览器需要支持 WebAssembly（现代浏览器均支持）
- 依赖：`modern-xlsx@^1.2.0` 随包自动安装（直接 dependency，wasm 二进制由本包转发暴露）；`xlsx`（SheetJS）为可选兜底依赖

## 安装

```bash
pnpm add @marcusok/excel-exporter
```

> **Node 版本说明**：`modern-xlsx` 声明了 `engines.node >= 24`，但其 WASM 核心面向浏览器，Node 22 实际可正常使用——本包在 Node 22 上开发并通过 CI。若包管理器在 Node 22 下因 engines 校验拒绝安装（如 pnpm 开启 `engine-strict` 时报错），可在项目 `.npmrc` 中加一行 `engine-strict=false`，或升级到 Node ≥ 24。

> `modern-xlsx` 是直接 `dependency`：本包锁定经过测试的引擎版本，并在构建时把 `modern-xlsx.wasm` 转发进自己的 `dist/` 通过 `exports` 暴露，保证二进制与 JS 胶水版本一致；同时 modern-xlsx 自身的 `exports` 未暴露 wasm 子路径，直接深导入会被打包器拒绝，由本包转发后才能 `?url` 引用。

需要兜底时额外安装 SheetJS。请使用官方 CDN 的 tarball 而非 npm 版本：npm 上的最后一个版本（`0.18.5`）已停止维护且存在已知 CVE（CVE-2023-30533、CVE-2024-22363），修复版只在 SheetJS 官方 CDN 发布。

```bash
pnpm add https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz
```

不安装时，兜底路径会自动从 SheetJS 官方 CDN 加载 `xlsx.mjs`（0.20.3），但生产环境更推荐自托管。

## 浏览器：静态资源部署

浏览器运行需要以下资源可被站点访问（两份文件都在本包 `dist/` 里，构建时 wasm 由本包自动转发）：

| 资源               | 何时需要                                                                                                                                              |
| ------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| `modern-xlsx.wasm` | 除显式 `mode: "stream"`（纯 JS 流式，无样式）外都需要（约 1.9MB，样式引擎核心），`configureWasm({ wasmUrl })` 指定                                    |
| `export.worker.js` | 仅进入 Worker 的路径需要，`configureWasm({ workerUrl })` 指定（auto ≥ 20,000 行，以及显式 `mode: "worker"` / `mode: "stream"`）；小数据量可完全不配置 |

### Vite（推荐）

用 `?url` 后缀导入资源，Vite 开发态直接服务、构建时自动 hash 拷贝进 `dist/assets/`——无需拷贝插件、无需 `public/` 目录：

```ts
// main.ts
import { configureWasm } from "@marcusok/excel-exporter";
import wasmUrl from "@marcusok/excel-exporter/dist/modern-xlsx.wasm?url";
// 可选：仅大表（≥ 2 万行）或显式 worker/stream 模式需要
import workerUrl from "@marcusok/excel-exporter/dist/export.worker.js?url";

configureWasm({ wasmUrl, workerUrl });
```

### 其他构建方案（拷贝文件）

打包器不支持资产 URL 导入时，在构建钩子里从 `require.resolve` 反推本包 `dist/` 真实路径，拷贝到静态目录（避免硬编码 node_modules 路径，pnpm 符号链接下更稳）：

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

应用入口统一配置：

```ts
import { configureWasm } from "@marcusok/excel-exporter";

configureWasm({
  wasmUrl: "/assets/modern-xlsx.wasm",
  workerUrl: "/assets/export.worker.js",
});
```

### configureWasm 参数

| 参数         | 类型            | 默认值   | 说明                                                            |
| ------------ | --------------- | -------- | --------------------------------------------------------------- |
| `wasmUrl`    | `string \| URL` | —        | 自托管 WASM 地址，生产强烈建议显式配置避免 CDN 漂移             |
| `workerUrl`  | `string \| URL` | —        | `export.worker.js` 地址，worker 模式必填                        |
| `timeoutMs`  | `number`        | `10_000` | 单次加载超时                                                    |
| `maxRetries` | `number`        | `3`      | 最大尝试次数；默认 3 次尝试，失败后按 300ms、600ms 指数退避等待 |

`configureWasm` 是合并语义：仅当 `wasmUrl` 变化时才重置已加载（或加载中）的 WASM 实例，只改超时/重试不会造成重复初始化；若此前加载失败（error 态），任意 `configureWasm` 调用都会清除错误态，下次导出按新配置重试。

## Node / SSR

Node 环境无需部署浏览器静态资源，也**无需任何初始化样板**：未配置 `wasmUrl` 时，引擎会在首次使用时自动通过 `createRequire` 从 `node_modules` 定位 `modern-xlsx.wasm` 并同步初始化（`initWasmSync`），随后即可正常导出带样式的文件。若想自行控制初始化时机（例如把一次性的同步读取+编译提前到进程启动期），或使用可 fetch 的 HTTP 地址，参见 [Node/SSR](/zh/packages/excel-exporter/guide/09-node-ssr)。`auto` 模式下 Node 不会走 Worker，而是主线程执行；≥ 5 万行自动切换流式路径（流式路径不依赖 WASM）。
