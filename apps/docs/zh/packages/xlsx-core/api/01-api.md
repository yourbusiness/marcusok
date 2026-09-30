# 加载器 API

`@marcusok/xlsx-core` 自身代码的完整接口面——即 loader。本包其余的导出都是 modern-xlsx 的再导出，列在[引擎再导出面](/zh/packages/xlsx-core/guide/02-engine-surface)。

```ts
import {
  WasmLoader,
  configureWasm,
  getWasmLoader,
  defaultWasmUrl,
  type LoaderOptions,
  type LoadState,
} from "@marcusok/xlsx-core";
```

## `configureWasm(options: LoaderOptions): void`

把 `options` 增量合并进共享 loader 的当前配置。完全可选：什么都不配时，资产解析到随包发布的位置。

- 此前的加载**错误**总是被清除，因此失败之后再调用它，下一次 `ensureLoaded()` 会用新配置重试，而不是永远抛"previously failed"。
- 只改 `timeoutMs` / `maxRetries` / `workerTimeoutMs` 会保留已加载的 WASM 模块。
- 改 `wasmUrl` 会把 `ready` / `loading` 的 loader 重置为 `idle`。这保证 `initWasm` 被**调用**时带着新 URL——但不保证新 URL 被**采用**：modern-xlsx 的 `initWasm` 是幂等的且只保留一个 in-flight promise，因此在已初始化 WASM 的线程上这次调用是静默空操作，仍在挂起的初始 fetch 也无法中止或改道。新 URL 只在全新 JS realm 中真正生效（刷新页面，或在共享 worker 被终止后新建 worker）。该限制适用时，`updateOptions` 会打印 console 警告。
- 改 `workerUrl` / `parseWorkerUrl` 一定会告警：共享 worker 只在创建时读取一次脚本 URL，之后一直复用。

## `getWasmLoader(): WasmLoader`

返回模块级单例。任何 @marcusok 包的导入方拿到的都是同一个对象。

## `defaultWasmUrl(): URL`

随包发布二进制的位置：`new URL("./modern-xlsx.wasm", import.meta.url)`，保留为字面量 `new URL` 表达式以便打包器把它发射成资产。当要把解析后的默认值转发给无法解析相对 URL 的地方时用它——worker 收到的是经 `postMessage` 传过去的 `String(wasmUrl ?? defaultWasmUrl())`，因为 `URL` 对象不可结构化克隆。

## `class WasmLoader`

| 成员                  | 签名                            | 说明                                                                                  |
| --------------------- | ------------------------------- | ------------------------------------------------------------------------------------- |
| `constructor`         | `(opts?: LoaderOptions)`        | 预置 `{ timeoutMs: 10_000, maxRetries: 3 }`，其余合并进来                             |
| `supported`（getter） | `boolean`                       | `typeof WebAssembly !== "undefined" && typeof WebAssembly.instantiate === "function"` |
| `isReady`（getter）   | `boolean`                       | `state === "ready"`——状态机唯一的公开视图，且不会触发加载                             |
| `getOptions()`        | `Readonly<LoaderOptions>`       | 当前合并后的配置（业务包就是从这里读 `workerTimeoutMs`）                              |
| `updateOptions(opts)` | `(opts: LoaderOptions) => void` | `configureWasm` 对单例调用的就是这个                                                  |
| `ensureLoaded()`      | `() => Promise<void>`           | 单飞（single-flight）；引擎可用时 resolve，否则以加载错误 reject                      |

自行构造实例是受支持的，但它隔离的**只有 loader 状态**——modern-xlsx 的初始化是模块全局的，同一线程上第二个实例变不出第二份引擎。

`ensureLoaded()` 在 ready 时立即 resolve，在 loading 时加入在途 promise，在 idle 时启动加载，在 error 态以 `[xlsx-core] WASM load previously failed; call configureWasm() to retry with new settings` reject。加载失败以 `[xlsx-core] WASM load failed after <n> attempts: <最后一次错误>` reject；环境不支持 WebAssembly 时在任何尝试之前就以 `[xlsx-core] WebAssembly not supported in this environment` reject。

## `LoaderOptions`

| 字段              | 类型            | 默认值                                 | 消费方                     | 含义                                                                                    |
| ----------------- | --------------- | -------------------------------------- | -------------------------- | --------------------------------------------------------------------------------------- |
| `wasmUrl`         | `string \| URL` | `defaultWasmUrl()`（随包发布的二进制） | loader                     | 显式 WASM 位置；同时会禁用 Node 的磁盘自动初始化                                        |
| `workerUrl`       | `string \| URL` | 导出包随包发布的 `export.worker.js`    | `@marcusok/excel-exporter` | 导出 worker 脚本 URL                                                                    |
| `parseWorkerUrl`  | `string \| URL` | 预览随包发布的 `parse.worker.js`       | `@marcusok/excel-preview`  | 解析 worker 脚本 URL——独立字段，使同页使用两个包时不会互拿对方的 worker                 |
| `timeoutMs`       | `number`        | `10_000`                               | loader                     | 单次加载超时                                                                            |
| `maxRetries`      | `number`        | `3`                                    | loader                     | 总尝试次数（含首次）；退避为 `300 * 2 ** (attempt - 1)` 毫秒                            |
| `workerTimeoutMs` | `number`        | `120_000`                              | 两个业务包                 | Worker 操作超时；超时会终止共享 worker 并拒绝其兄弟请求。必须 `> 0`——不存在"禁用"的取值 |

所有字段均可选。只有 loader 侧默认值存放在 loader 里；worker 相关默认值由消费包解析（见 [WASM 加载器指南](/zh/packages/xlsx-core/guide/01-loader#loaderoptions)）。

## `LoadState`

```ts
type LoadState = "idle" | "loading" | "ready" | "error";
```

| 取值      | 含义                                         | `ensureLoaded()` 行为                        |
| --------- | -------------------------------------------- | -------------------------------------------- |
| `idle`    | 尚未开始，或 loader 因 `wasmUrl` 变更被重置  | 启动加载                                     |
| `loading` | 有加载在途                                   | 加入同一个 promise                           |
| `ready`   | 引擎已在本线程初始化                         | 立即 resolve                                 |
| `error`   | 所有尝试均失败（迟到的成功会自愈回 `ready`） | 持续 reject，直到 `configureWasm()` 清除该态 |

状态字段本身是私有的：请通过 `isReady` 与 `ensureLoaded()` 的结果观察。

## `@marcusok/xlsx-core/tsup`

```ts
import {
  rewriteWasmBgUrl,
  dropNodeFsPromises,
  type EsbuildPlugin,
} from "@marcusok/xlsx-core/tsup";

declare const rewriteWasmBgUrl: EsbuildPlugin;
declare const dropNodeFsPromises: EsbuildPlugin;
```

`EsbuildPlugin` 即 `NonNullable<Options["esbuildPlugins"]>[number]`——从 tsup 自身的选项形状推导而来，两个插件可直接放进 `esbuildPlugins`。仅限 Node（内部 import `node:fs`）；用法见[在包中接入引擎层](/zh/packages/xlsx-core/guide/03-package-integration)。

## 完整示例

```ts
import {
  configureWasm,
  defaultWasmUrl,
  getWasmLoader,
  readBuffer,
  formatCellRich,
  ModernXlsxError,
  WASM_INIT_FAILED,
  type LoaderOptions,
} from "@marcusok/xlsx-core";

// 1. 可选：自托管资产，以及/或放宽首次加载超时
configureWasm({
  wasmUrl: "/assets/modern-xlsx.wasm",
  timeoutMs: 15_000,
} satisfies LoaderOptions);

// 2. 可选：提前预热引擎（Node 下即同步从磁盘读取）
const loader = getWasmLoader();
if (loader.supported && !loader.isReady) {
  await loader.ensureLoaded(); // 失败会抛错——configureWasm() 可重新武装它
}

// 3. 使用引擎
try {
  const workbook = await readBuffer(bytes);
  console.log(workbook.sheetNames, formatCellRich(1234.5, "#,##0.00").text);
} catch (err) {
  if (err instanceof ModernXlsxError && err.code === WASM_INIT_FAILED) {
    console.error(
      "引擎不可用，实际解析到的 wasm URL 为",
      String(defaultWasmUrl()),
    );
  }
}
```
