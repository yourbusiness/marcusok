# WASM 加载器

loader 是这个包存在的理由。它负责定位 WASM 二进制、保证每个 JS realm 只初始化一次，并用一个类型化的状态机上报失败——浏览器与 Node 两侧都零配置。

```ts
import { getWasmLoader, readBuffer } from "@marcusok/xlsx-core";

// 可选：仅在自托管副本 / CDN / 自定义构建时需要
// configureWasm({ wasmUrl: "/assets/modern-xlsx.wasm" });

await getWasmLoader().ensureLoaded(); // 已 ready 时是空操作
const workbook = await readBuffer(bytes);
```

多数消费方不会显式调用 `ensureLoaded()`：业务包在首次操作时替它 await，引擎因此是懒初始化的。当你想把一次性的读取+编译记在某个明确时刻（进程启动、加载动画）而不是第一次用户操作上时，再自己调用它。

## 默认定位机制

| 环境   | 行为                                                                                                                                                                                                                                                                                                                                                                                                                  |
| ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 浏览器 | `defaultWasmUrl()` 返回 `new URL("./modern-xlsx.wasm", import.meta.url)`——正是 Vite 的 `vite:asset-import-meta-url` 与 webpack 5 会静态改写、把文件发射为你产物旁 hash 资产的那种写法。随后 loader fetch 该 URL 并调用引擎的 `initWasm(wasmUrl)`。                                                                                                                                                                    |
| Node   | 未配置任何东西时，loader 从磁盘读取自己的 `dist/modern-xlsx.wasm` 并调用 `initWasmSync(bytes)`——Node 的 `fetch` 会拒绝浏览器默认值产生的 `file://` URL，同步路径是唯一正确的选择。`node:fs` 用计算式说明符动态导入，使本模块对浏览器目标也是打包安全的。查找链：发布入口旁的 `./modern-xlsx.wasm`（常规情形）→ `../dist/modern-xlsx.wasm`（本模块从 `src/` 运行时的位置，即仓库测试与源码别名化的 monorepo 消费方）。 |

一旦你显式配置了 `wasmUrl`，Node 自动初始化即被跳过——你要的就是那个 URL，loader 会像浏览器那样经 `initWasm` 去取。自动初始化过程中的任何定位/读取/初始化失败都会静默返回 `false` 并落到常规 `initWasm` 路径，所以"打包了代码但没带上资产"的部署形态，降级表现与自动初始化存在之前完全一致。

## LoaderOptions

每个字段都可选。`configureWasm(options)` 是增量合并。

| 参数              | 类型            | 默认值                                           | 含义                                                                                                                                                                                                                 |
| ----------------- | --------------- | ------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `wasmUrl`         | `string \| URL` | 随包发布的 `dist/modern-xlsx.wasm`               | 覆盖为自托管副本、CDN，或不支持资产 URL 的打包器                                                                                                                                                                     |
| `workerUrl`       | `string \| URL` | `@marcusok/excel-exporter` 随包发布的导出 worker | 导出 worker 脚本。只由导出包读取——单字段无法同时服务两个包（预览会拉起导出 worker），故有下面这个独立选项                                                                                                            |
| `parseWorkerUrl`  | `string \| URL` | `@marcusok/excel-preview` 随包发布的解析 worker  | 解析 worker 脚本。只由预览读取                                                                                                                                                                                       |
| `timeoutMs`       | `number`        | `10_000`                                         | 单次加载超时                                                                                                                                                                                                         |
| `maxRetries`      | `number`        | `3`                                              | 总尝试次数（含首次）。失败后按 300ms、600ms 退避                                                                                                                                                                     |
| `workerTimeoutMs` | `number`        | `120_000`                                        | Worker 操作超时，由两个业务包消费。超时操作会终止共享 worker 并拒绝其兄弟请求，故只应针对确实巨大的负载调高。必须 `> 0`：两个包都把该值原样交给 `setTimeout`（`value ?? 120_000`），`0` 的含义是"立即超时"而非"禁用" |

默认值来自两处，读表时值得知道这一点：`timeoutMs` / `maxRetries` 由 `new WasmLoader()` 自身写入（构造函数预置 `{ timeoutMs: 10_000, maxRetries: 3 }`），`wasmUrl` 则是在加载时回退到 `defaultWasmUrl()`（`this.opts.wasmUrl ?? defaultWasmUrl()`）。worker 两个选项正相反：loader 从不为它们解析默认值，只存你传入的内容，各业务包各自取自己随包发布的默认值——导出包用 `workerUrl ?? new URL("./export.worker.js", import.meta.url)`，预览用 `parseWorkerUrl ?? new URL("./parse.worker.js", import.meta.url)`——并在调用时读取 `getOptions().workerTimeoutMs ?? 120_000`。因此，直接消费本包的项目根本没有 worker 脚本：这两个选项是为两个业务包存在的。

## 超时与重试语义

一次尝试 = 一次 `initWasm(wasmUrl)` 调用与一个 `timeoutMs` 计时器的竞速。失败后下一次尝试等待 `300 * 2 ** (attempt - 1)` 毫秒——默认三次尝试即 300ms、600ms。全部尝试失败时 `ensureLoaded()` 以如下信息 reject：

```
[xlsx-core] WASM load failed after 3 attempts: <最后一次错误信息>
```

有两种失败形态值得区分：

- **环境不支持 WebAssembly。** `loadWithRetry` 立即以 `[xlsx-core] WebAssembly not supported in this environment` reject，不消耗重试次数。在提供依赖 WASM 的界面之前，先查 `getWasmLoader().supported`（或 `typeof WebAssembly`）。
- **所有尝试都超时了，但底层加载随后成功。** modern-xlsx 的 `initWasm` 在模块级保留单一的 in-flight promise，所以慢网络下那次原始 fetch 可能在 loader 放弃之后才 resolve。loader 会盯着这一点并自愈：迟到的成功把 `error` 拨回 `ready` 并清空挂起的 promise，此后无需任何 `configureWasm()` 调用即可正常工作。没有这个挂钩，`error` 态就会与引擎的真实状态永久矛盾。同一机制也解释了[导出包安装页](/zh/packages/excel-exporter/guide/02-installation#vite-开发服务器-预构建注意事项)上的那个诊断陷阱：首次失败之后，后续尝试报的是 `WASM load previously failed` 而不是原始原因。

## 加载状态机

`LoadState` 为 `"idle" | "loading" | "ready" | "error"`，而它是**私有状态**：该字段不属于公开类成员，请通过两个 getter 与 `ensureLoaded()` 的结果来观察。

| 状态      | 如何进入                                           | `ensureLoaded()` 的行为                                                                                     |
| --------- | -------------------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `idle`    | 初始状态；`configureWasm` 重置 loader 后也回到此态 | 置为 `loading` 并启动重试循环                                                                               |
| `loading` | 有加载正在进行                                     | 返回**同一个** promise——并发调用方会单飞（single-flight）到同一次加载，而不是再起一次                       |
| `ready`   | 所有尝试完成（或迟到成功自愈）                     | 立即返回；`loader.isReady` 为 `true`                                                                        |
| `error`   | 所有尝试均失败                                     | 抛出 `[xlsx-core] WASM load previously failed; call configureWasm() to retry with new settings`，自身不重试 |

何时查它：需要**不触发工作**的判断（"能展示带样式的路径吗？"）用 `isReady`；需要"把引擎给我"的判断用 `ensureLoaded()`。状态只在 `ensureLoaded()` 内写入，并以 promise 身份校验守护，因此被取代的旧加载（加载途中 URL 变更）既不能把 loader 标成 ready，也盖不掉新加载的结果。

`configureWasm()` 无条件清除 `error` 态；而 `wasmUrl` 发生**变更**时，还会把 `ready` / `loading` 重置回 `idle`。只改超时/重试则保留已加载的模块——没有理由因为超时值变了就重新编译 WASM。

## WasmLoader 与 defaultWasmUrl()

- `WasmLoader` 是单例背后的类。`getWasmLoader()` 返回模块级实例；自行 `new WasmLoader({ timeoutMs: 1000 })` 也受支持（测试、隔离 loader 状态），但隔离的**只是 loader 状态**——modern-xlsx 的 `initWasm` 是模块全局的，同一线程上第二个 loader 也变不出第二份引擎实例。
- `defaultWasmUrl()` 以 `URL` 对象返回随包发布二进制的位置。它之所以公开，是因为消费方需要**转发**解析后的默认值：被复制或改名的 worker 脚本无法可靠地按相对路径解析 `./modern-xlsx.wasm`，所以业务包改为经 `postMessage` 传 `String(wasmUrl ?? defaultWasmUrl())`（`URL` 对象不可结构化克隆，会抛 `DataCloneError`）。

## 资产自托管与 CDN 覆盖

```ts
import { configureWasm } from "@marcusok/xlsx-core";

configureWasm({ wasmUrl: "https://cdn.example.com/modern-xlsx.wasm" });
```

也可以把随包文件接到打包器的资产导入上：

```ts
import wasmUrl from "@marcusok/xlsx-core/dist/modern-xlsx.wasm?url";
configureWasm({ wasmUrl });
```

**要在首次加载之前调用。** modern-xlsx 的 `initWasm` 带"首次成功者胜"的幂等守卫、且只保留一个 in-flight promise，因此成功加载后改 `wasmUrl` 不会重新加载，初始 fetch 挂起期间改也无法改道那次 fetch。新 URL 只在全新 JS realm 中生效——刷新页面，或在共享 worker 被终止后新建 worker。只要该限制适用，`updateOptions()` 就会打印 console 警告（`workerUrl` / `parseWorkerUrl` 同理：共享 worker 只在创建时读一次脚本 URL）。

worker 脚本还有第二重约束：浏览器在 `Worker` 构造时拒绝跨域脚本，可走裸 CDN 的只有 WASM 二进制。两个业务包各自的覆盖方式见下方[从其他包使用同一个 loader](#从其他包使用同一个-loader)，以及[导出包安装指南](/zh/packages/excel-exporter/guide/02-installation#可选-configurewasm)与[预览资产指南](/zh/packages/excel-preview/guide/02-assets)。

## Vite 开发服务器：预构建注意事项

导出包文档里那条注意事项对直接消费本包的项目同样成立，成因与修复都相同：Vite 开发服务器把依赖预构建进 `node_modules/.vite/deps/`，那里的 `import.meta.url` 已不再指向真实的 `dist/`，默认 WASM URL 于是解析到一个不存在的路径，请求被 `index.html` 应答。把提供 loader 的包加进 `optimizeDeps.exclude`：

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ["@marcusok/xlsx-core"] },
});
```

`vite build` 不受影响（生产管线会把 wasm 正确发射为 hash 资产），上文 `?url` + `configureWasm` 的接法也不用动 `optimizeDeps`。完整排查——包括 `expected magic word` 编译错误，以及业务包表现出的静默降级症状——见[导出包安装说明](/zh/packages/excel-exporter/guide/02-installation#vite-开发服务器-预构建注意事项)；Node 服务端另有其形态，见 [Node/SSR 指南](/zh/packages/excel-exporter/guide/09-node-ssr)。

## 从其他包使用同一个 loader

`configureWasm` 与 `getWasmLoader` 被两个业务包再导出，三个名字指向的都是本包这一个 loader：

```ts
import { configureWasm, getWasmLoader } from "@marcusok/excel-exporter"; // 或 @marcusok/excel-preview
```

如果你已经在用其中之一，这就是推荐的入口——本文记载的选项语义、默认值与警告完全一致；[预览的 API 参考](/zh/packages/excel-preview/api/01-create-preview#其他导出)也把它们列进了自己的导出面，原因相同。
