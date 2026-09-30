# 资产与自托管

预览器带两个运行时资产：WASM 引擎二进制（1.9MB，gzip 约 650KB）与自包含解析 worker。两者在打包器与 Node 中都自动定位，只有自托管场景需要配置。

## 默认定位机制

| 资产        | 默认位置                                        | 机制                                                                                |
| ----------- | ----------------------------------------------- | ----------------------------------------------------------------------------------- |
| WASM 二进制 | `@marcusok/excel-preview/dist/modern-xlsx.wasm` | `new URL(<file>, import.meta.url)`——打包器产出 hash 资产；Node 经 node_modules 定位 |
| 解析 worker | `@marcusok/excel-preview/dist/parse.worker.js`  | 同一模式；worker 是单文件自包含 ESM（没有会 404 的兄弟导入）                        |

浏览器里主线程完全不需要 WASM：解析在 worker 内（worker 自带一份引擎），主线程唯一的引擎调用是格式化——纯 JS。Node（无 Worker）下主线程从磁盘同步初始化 WASM，零样板。

引擎在构建期打包进本包 `dist`（2.0 之前曾以独立的 `xlsx-core` 包分发）。同页使用 `excel-exporter` 与 `excel-preview` 时主线程各持一份引擎副本——两份二进制完全相同，通常被内容 hash 命名与 HTTP 缓存去重成一份资产。各包的自包含 worker 因浏览器 module worker 无法解析裸导入，必然各带一份——这是与导出包一致的既有事实。

## 自托管 / CDN

两个 URL 指到你的副本（调用形态与导出包一致——各包配置各自的 loader）：

```ts
import { configureWasm } from "@marcusok/excel-preview";

configureWasm({
  wasmUrl: "https://cdn.example.com/modern-xlsx.wasm",
  parseWorkerUrl: "https://cdn.example.com/parse.worker.js",
});
```

注意字段名：解析 worker 走独立的 `parseWorkerUrl` 选项——loader 源码由各 @marcusok 包共用，各包的 worker 是不同脚本，单字段会让两包互拿对方的 worker（`workerUrl` 是**导出** worker 的选项，由 @marcusok/excel-exporter 读取）。

> **worker 脚本只能同源。** 浏览器在 `Worker` 构造时直接拒绝跨域 worker 脚本（抛 `SecurityError`）——`parseWorkerUrl` 填裸 CDN 地址是加载不出来的。预览会检测到该情形、在控制台告警并回退主线程解析，但要想真正用上 worker，需从自己的源提供该文件（本地副本或反向代理）；可走裸 CDN 的只有 WASM 二进制。

配合打包器资产导入：

```ts
import wasmUrl from "@marcusok/excel-preview/dist/modern-xlsx.wasm?url";
import parseWorkerUrl from "@marcusok/excel-preview/dist/parse.worker.js?url";
configureWasm({ wasmUrl, parseWorkerUrl });
```

从导出包 loader 继承的注意事项（两者共用）：

- `configureWasm` 要在**第一次预览之前**调用；成功加载后改 WASM URL 只在全新 JS realm 生效（刷新页面或新建 worker）。
- worker 加载失败时自动回退主线程解析——预览仍可用，代价是大文件解析期间阻塞主线程。

## loader 默认值

共享 loader 的可调项与默认值（与导出包一致——同一个 loader、同一组配置）：

| 配置项            | 默认值    | 含义                                                                                                                                                 |
| ----------------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| `timeoutMs`       | `10_000`  | 单次尝试的 WASM 加载超时（fetch/instantiate）                                                                                                        |
| `maxRetries`      | `3`       | 总加载尝试次数（含首次；重试间隔 300ms / 600ms 退避）                                                                                                |
| `workerTimeoutMs` | `120_000` | 解析 worker 的操作超时。超时的解析会终止共享 worker、拒绝在途请求且不在主线程重跑；下次解析重建新 worker。必须 > 0——`0`/负数是"立即超时"，不是"禁用" |

## Vite dev server 注意事项

与导出包相同：若 WASM 请求返回 HTML（`content-type: text/html`，编译错误含 `expected magic word`），把包加进 `optimizeDeps.exclude`，阻止依赖预打包拦截资产。完整排查思路见[导出包安装文档](/zh/packages/excel-exporter/guide/02-installation#vite-开发服务器-预构建注意事项)。
