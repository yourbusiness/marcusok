# 资产与自托管

预览器带两个运行时资产：WASM 引擎二进制（约 2MB，gzip 约 650KB）与自包含解析 worker。两者在打包器与 Node 中都自动定位，只有自托管场景需要配置。

## 默认定位机制

| 资产        | 默认位置                                       | 机制                                                                                |
| ----------- | ---------------------------------------------- | ----------------------------------------------------------------------------------- |
| WASM 二进制 | `@marcusok/xlsx-core/dist/modern-xlsx.wasm`    | `new URL(<file>, import.meta.url)`——打包器产出 hash 资产；Node 经 node_modules 定位 |
| 解析 worker | `@marcusok/excel-preview/dist/parse.worker.js` | 同一模式；worker 是单文件自包含 ESM（没有会 404 的兄弟导入）                        |

浏览器里主线程完全不需要 WASM：解析在 worker 内（worker 自带一份引擎），主线程唯一的引擎调用是格式化——纯 JS。Node（无 Worker）下主线程从磁盘同步初始化 WASM，零样板。

引擎位于共享包 [@marcusok/xlsx-core](https://www.npmjs.com/package/@marcusok/xlsx-core)：同页使用 `excel-exporter` 与 `excel-preview` 时主线程只加载一份引擎与一份 WASM。各包的自包含 worker 因浏览器 module worker 无法解析裸导入，必然各带一份——这是与导出包一致的既有事实。

## 自托管 / CDN

两个 URL 指到你的副本（与导出包同一个调用——loader 是共享的）：

```ts
import { configureWasm } from "@marcusok/excel-preview";

configureWasm({
  wasmUrl: "https://cdn.example.com/modern-xlsx.wasm",
  parseWorkerUrl: "https://cdn.example.com/parse.worker.js",
});
```

注意字段名：解析 worker 走独立的 `parseWorkerUrl` 选项——共享 loader 服务同页多个 @marcusok 包，各包的 worker 是不同脚本，单字段会让两包互拿对方的 worker（`workerUrl` 是**导出** worker 的选项，由 @marcusok/excel-exporter 读取）。

> **worker 脚本只能同源。** 浏览器在 `Worker` 构造时直接拒绝跨域 worker 脚本（抛 `SecurityError`）——`parseWorkerUrl` 填裸 CDN 地址是加载不出来的。预览会检测到该情形、在控制台告警并回退主线程解析，但要想真正用上 worker，需从自己的源提供该文件（本地副本或反向代理）；可走裸 CDN 的只有 WASM 二进制。

配合打包器资产导入：

```ts
import wasmUrl from "@marcusok/xlsx-core/dist/modern-xlsx.wasm?url";
import parseWorkerUrl from "@marcusok/excel-preview/dist/parse.worker.js?url";
configureWasm({ wasmUrl, parseWorkerUrl });
```

从导出包 loader 继承的注意事项（两者共用）：

- `configureWasm` 要在**第一次预览之前**调用；成功加载后改 WASM URL 只在全新 JS realm 生效（刷新页面或新建 worker）。
- worker 加载失败时自动回退主线程解析——预览仍可用，代价是大文件解析期间阻塞主线程。

## Vite dev server 注意事项

与导出包相同：若 WASM 请求返回 HTML（`content-type: text/html`，编译错误含 `expected magic word`），把包加进 `optimizeDeps.exclude`，阻止依赖预打包拦截资产。完整排查思路见[导出包安装文档](/zh/packages/excel-exporter/guide/02-installation#vite-dev-server-pre-bundling-caveat)。
