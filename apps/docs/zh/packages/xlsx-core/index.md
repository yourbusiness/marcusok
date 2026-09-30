# @marcusok/xlsx-core

全仓围绕 [modern-xlsx](https://github.com/ABCrimson/modern-xlsx)（Rust + WASM）的唯一集成点：WASM 加载、资产分发，以及一层稳定的引擎再导出面。除了 loader 之外它自身不含任何业务逻辑——其余导出的都是引擎的公开 API，只是由本包转手发布，好让各 @marcusok 包无需直接依赖 `modern-xlsx`。

引擎运行时在构建期被打包**进**本包的 `dist/`，因此消费方看不到任何外部运行时依赖，也不受 modern-xlsx 自身 `engines.node >= 24` 声明的影响。`package.json` 里钉住的 `modern-xlsx` 依赖只为让 TypeScript 消费方能解析再导出的类型（见[给下游包作者](/zh/packages/xlsx-core/guide/03-package-integration)）。

## 包含什么

| 层次        | 提供的能力                                                                                                                                                           |
| ----------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| WASM 加载器 | `configureWasm()` / `getWasmLoader()` / `WasmLoader` / `defaultWasmUrl()`——浏览器 fetch 与 Node 同步初始化双路径、单次加载超时与重试、全页共用一个 loader 实例       |
| 资产分发    | `dist/modern-xlsx.wasm`（约 1.9MB），通过 exports 映射再发布，`@marcusok/xlsx-core/dist/modern-xlsx.wasm?url` 在 Vite / webpack 5 中可解析                           |
| 引擎再导出  | 各 @marcusok 包共用的 modern-xlsx 运行时面：`Workbook`、`readBuffer`、引用/格式化/日期工具、类型化错误（[完整清单](/zh/packages/xlsx-core/guide/02-engine-surface)） |
| 构建插件    | `@marcusok/xlsx-core/tsup`——下游包打包时复用的两个 esbuild 插件（[详见](/zh/packages/xlsx-core/guide/03-package-integration)）                                       |

## 何时直接使用

- **你在自研渲染器、写入器或数据管线。** 导出包与预览各自把数据整理成自己的模型；如果你想要的是引擎本体加它的原语——`readBuffer` 读出 `Workbook`、`formatCellRich`、`decodeRange`、`serialToDate`——这一层就是基建。
- **你需要类型化的错误面。** `ModernXlsxError` 带机器可读的 `code`；本包再导出的四个常量覆盖了业务包用来分流的格式/密码/初始化三类情形（见[引擎再导出面](/zh/packages/xlsx-core/guide/02-engine-surface#类型化错误)）。
- **你的目标不是 Excel。** 引擎读写 OOXML；任何只需要读取侧的用途（差异比对、校验、导入管线）都能用它，而不必带上导出或预览的 UI。
- **你希望全页只有一份引擎实例。** 从这里导入 loader（或在这里配置它），拿到的就是业务包用的那个单例——同页即使还用着业务包，也依然只加载一份 WASM 模块。

## 何时不该直接使用

要导出工作簿？用 [@marcusok/excel-exporter](/zh/packages/excel-exporter/)。要预览？用 [@marcusok/excel-preview](/zh/packages/excel-preview/)。两者都**再导出**了 `configureWasm` / `getWasmLoader`（它们就是 `@marcusok/xlsx-core` 自己的函数，被转手发布），所以从任一包导入、配置的都是同一个共享 loader——仅为自托管资产而新增本依赖没有意义。

## 安装

```bash
pnpm add @marcusok/xlsx-core
```

环境要求：Node `>= 22`（本包的 `engines`），浏览器需支持 WebAssembly。不会引入任何运行时依赖——引擎就装在本包的 `dist/` 里。

## exports 映射

| 子路径                                      | 解析到                              | 用途                                                                            |
| ------------------------------------------- | ----------------------------------- | ------------------------------------------------------------------------------- |
| `@marcusok/xlsx-core`                       | `dist/index.js` + `dist/index.d.ts` | loader 与引擎再导出面                                                           |
| `@marcusok/xlsx-core/tsup`                  | `dist/tsup.js` + `dist/tsup.d.ts`   | `rewriteWasmBgUrl` / `dropNodeFsPromises`，供下游包自己的 `tsup.config.ts` 使用 |
| `@marcusok/xlsx-core/dist/modern-xlsx.wasm` | WASM 二进制                         | 支持资产导入的打包器用 `?url` 导入（`...modern-xlsx.wasm?url`）                 |
| `@marcusok/xlsx-core/package.json`          | 包清单                              | 需要读取版本 / engines 范围、而不想导入模块的工具链                             |

## 与两个业务包的关系

- **同一个 loader。** 对任何导入方，`getWasmLoader()` 返回的都是同一个 `WasmLoader` 对象：通过 `@marcusok/excel-exporter` 或 `@marcusok/excel-preview` 配置 WASM URL，配置的就是本包持有的这一个。
- **主线程同一份 WASM 二进制。** 两个业务包在主入口构建中都把 `@marcusok/xlsx-core` 保持为 external，因此同页使用多个 @marcusok 包时，引擎只编译、实例化一次。
- **worker 是例外。** 各包的自包含 worker 会各带一份引擎——浏览器 module worker 无法解析裸导入，这份重复是固有代价（导出包与预览各自的 worker 路径见它们自己的文档）。

两种环境下配置都是可选的：资产默认解析到随包发布的位置。先读 [WASM 加载器指南](/zh/packages/xlsx-core/guide/01-loader)；同一套接口的精简版在 [API 参考](/zh/packages/xlsx-core/api/01-api)。
