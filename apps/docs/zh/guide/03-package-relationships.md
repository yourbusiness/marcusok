# 包之间的关系与选型

生态对外发布五个包、分属两个大类（大类定义见[生态](/zh/guide/)）。本页回答选型时常见的问题：各包负责什么、各自内置了什么、多包能否同页共存、版本如何对应。

## 五个发布包

| 包                                                         | 大类 | 职责                                                                         | 运行时依赖                                       |
| ---------------------------------------------------------- | ---- | ---------------------------------------------------------------------------- | ------------------------------------------------ |
| [`@marcusok/excel-exporter`](/zh/packages/excel-exporter/) | 导出 | 业务数据 → `.xlsx` 下载：样式、格式化、worker/流式路由与多级降级兜底         | 无（引擎本体已打包）                             |
| [`@marcusok/excel-preview`](/zh/packages/excel-preview/)   | 预览 | 只读渲染 `.xlsx` / `.xlsm` / `.csv`：worker 解析 + 虚拟滚动                  | 无（引擎本体已打包）                             |
| [`@marcusok/word-preview`](/zh/packages/word-preview/)     | 预览 | 只读渲染 `.docx` / `.docm`（docx-preview 引擎），带嗅探与加固默认值          | 无（docx-preview 与 jszip 已打包）               |
| [`@marcusok/pdf-preview`](/zh/packages/pdf-preview/)       | 预览 | 只读渲染 `.pdf`（pdf.js 引擎）：懒加载连续滚动、文本层可选中复制             | 无（pdf.js 已打包；worker/cmaps 随 `dist` 分发） |
| [`@marcusok/ppt-preview`](/zh/packages/ppt-preview/)       | 预览 | 只读渲染 `.pptx`（Rust/WASM 引擎）：懒加载连续滚动、0-based 页导航、加密解密 | 无（引擎与 wasm 已打包）                         |

每个包在运行期自包含。两个 Excel 包脚下有两个能力层——modern-xlsx 引擎层（仓库内部名 `xlsx-core`：WASM 加载器 + 稳定再导出面）与进度遮罩 UI（内部名 `progress-overlay`）——它们是**仓库私有包**：构建期打进 Excel 业务包的 `dist`（引擎层打进两个 Excel 包，遮罩只打进导出包，见下表），不发布到 npm。word、PDF 与 PPTX 预览以同样方式内置自己的引擎（docx-preview + jszip；pdf.js；Rust/WASM OOXML 引擎），运行时依赖为**零**。你只需要安装要用的那个包，也没有共享包的版本要对齐。

三点值得记住：

- **五个包是兄弟，不是上下层。** 互不依赖，可单独安装任意一个。Excel 引擎的任何概念都不会泄漏进预览包的 API，预览也不是为导出服务的。
- **每个 Excel 包各自内置整份 WASM 引擎。** modern-xlsx 的 JS 胶水与 WASM 二进制随各 Excel 包 `dist` 分发（也随各包的 worker 分发，见下文）。发布后的 manifest 没有 modern-xlsx 运行时依赖：仓库中保留的钉版声明只用于构建期的引擎/类型解析，发布出的 `.d.ts` 已自包含。word、PDF 与 PPTX 预览同样没有运行时依赖。
- **同页任意组合都可行**——见下文。

### 各自什么时候用

- **`@marcusok/excel-exporter`** —— 应用里有数据（表格、数组、ECharts 配置），需要产出给用户下载的文件。从[快速上手](/zh/guide/01-getting-started)开始；便捷适配器见[表格与 ECharts](/zh/packages/excel-exporter/api/05-table-and-echarts)。
- **`@marcusok/excel-preview`** —— 用户上传了 `.xlsx`，需要只读**展示**（不可编辑）。从[快速上手](/zh/packages/excel-preview/guide/01-quick-start)开始。
- **`@marcusok/word-preview`** —— 用户上传了 `.docx`，需要只读**展示**（不可编辑）。从[快速上手](/zh/packages/word-preview/guide/01-quick-start)开始。
- **`@marcusok/pdf-preview`** —— 用户上传了 `.pdf`，需要只读**展示**（不可编辑）。从[快速上手](/zh/packages/pdf-preview/guide/01-quick-start)开始。
- **`@marcusok/ppt-preview`** —— 用户上传了 `.pptx`，需要只读**展示**（不可编辑）。从[快速上手](/zh/packages/ppt-preview/guide/01-quick-start)开始。

## 多个包能同页共存吗

能——它们互不感知。唯一有实际代价的组合是两个 Excel 包同页：由于各自内置一份引擎，同页**同时**使用两者时，主线程会加载两份引擎实例。实践层面的含义：

- **网络成本通常会被去重。** 两个 Excel 包分发的 WASM 二进制完全相同；按内容 hash 命名资产的打包器（Vite、webpack 5）对相同内容只产出一个文件，浏览器 HTTP 缓存兜底。
- **真正的代价在内存。** 每份引擎副本各自编译、实例化自己的 WASM 模块（约 1.9 MB 二进制及其线性内存）。典型的管理后台页面只用其中一个包，通常遇不到；若确有性能敏感的页面同时跑两个 Excel 包，请把这一点计入。
- **worker 本来就是这样。** 浏览器 module worker 不能解析裸导入（import map 不作用于 `WorkerGlobalScope`），因此每个 worker 脚本从来都是自包含单文件、各带一份引擎——`export.worker.js`（导出包）与 `parse.worker.js`（预览包）各带一份，此前如此，现在不变。PDF 预览的 worker（`pdf.worker.min.mjs`，上游官方自包含构建）与这些互不相干。
- **word 预览不增加以上任何成本。** 无 WASM、无 worker——纯 JS/DOM 且实例级样式隔离，与任一或两个 Excel 包组合的代价只有它自身的包体积。PDF 预览每实例一个 pdf.js worker（不配 `wasmUrl` 就没有 WASM；cmaps/字体是经 HTTP 拉取的纯数据文件、从不编译），组合同样无碍。PPTX 预览每实例一个内联解析 worker + 一份约 2 MB WASM（排版与绘制在主线程），组合成本同上述 Excel 引擎——性能敏感页面控制实例数量。

### 哪些是各包自带的，哪些仍然分离

| 事物               | 各包自带还是共有？                                                                                                                                                                                                           |
| ------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 主线程引擎与 WASM  | **各 Excel 包自带**——各 Excel 包 `dist` 各持一份引擎副本与各自的 `modern-xlsx.wasm`；word 预览内置 docx-preview + jszip（无 WASM）                                                                                           |
| WASM loader 与配置 | **各包独立**——从 `@marcusok/excel-exporter` 导入的 `configureWasm` / `getWasmLoader` 作用于导出包副本；预览包同理                                                                                                            |
| worker 脚本        | **各包自带**——`export.worker.js`（导出包）与 `parse.worker.js`（预览包）是不同文件，各随所属包分发；word 预览没有 worker，PDF 预览分发上游的 `pdf.worker.min.mjs`，PPTX 预览的解析 worker 经 Blob 内联（无外部 worker 文件） |
| worker URL 选项    | **按包分字段**——`workerUrl` 属导出 worker，`parseWorkerUrl` 属解析 worker，单字段无法同时服务两者                                                                                                                            |
| 遮罩 UI            | **仅导出包**——预览包不打包、不使用它                                                                                                                                                                                         |

loader 一行是两个 Excel 包同页时唯一的行为差异点：**一次 `configureWasm` 只配置一个包的 loader，不会同时作用于两包。** 若在同页使用两个 Excel 包且自托管资产，请对每个用到的 Excel 包各调一次 `configureWasm`，传入相同的 `wasmUrl`：

```ts
import { configureWasm as configureExportWasm } from "@marcusok/excel-exporter";
import { configureWasm as configurePreviewWasm } from "@marcusok/excel-preview";

configureExportWasm({ wasmUrl, workerUrl }); // 导出包的 loader + export worker
configurePreviewWasm({ wasmUrl, parseWorkerUrl }); // 预览包的 loader + parse worker
```

超时参数（`timeoutMs`、`maxRetries`、`workerTimeoutMs`）挂在各自的 loader 上，作用于该包的资产加载与 worker 派发。降级同样按包独立：导出包降级为无样式快速流，预览包回退主线程解析；一个包降级不会改变另一个包的路径。

## Excel 包内的引擎层

引擎集成——WASM 加载（浏览器 fetch / Node 同步初始化、超时、重试、状态机）、资产分发与对 modern-xlsx 的稳定再导出面——收敛在一个内部层，两个 Excel 业务包各自打包它。对你的意义：

- **无需关心引擎版本。** 没有引擎包要安装、要锁定：引擎升级以你所用包的新版本形式到达。
- **`configureWasm` / `getWasmLoader` 由两个 Excel 包同签名再导出**——从你已在用的包导入即可，自托管资产不需要为此新增依赖。
- **WASM 二进制在两个 Excel 包各自的 `exports` 映射下分发**（`@marcusok/excel-exporter/dist/modern-xlsx.wasm`、`@marcusok/excel-preview/dist/modern-xlsx.wasm`），`?url` 导入零插件可解析。Node 从磁盘上已安装包的旁边定位二进制并同步初始化——零样板。

> 在单个包内，"loader 单例"是模块实例层面的保证：`getWasmLoader()` 对该包始终返回同一对象。跨两个 Excel 包则 loader 天然分离（各自打包一份）。

## 内置于导出包的遮罩

- **2.8.0 起**，每次导出默认显示遮罩。无需额外 import、无需包装函数。
- `overlay` 字段声明在 `ExportOptions` 上，三种取值——缺省/`true`（默认导出文案）、`false`（完全关闭：不挂载、不多让一帧）、或选项对象（定制文案、`delayMs`、主题等）。三者及自定义文案的合并语义见[进度遮罩](/zh/packages/excel-exporter/guide/11-overlay#取值)，本页不再重复。
- **该字段在全部三个入口都生效。** 3.0.2 起 `TableExportOptions` 与 `EChartsExportOptions` 也声明了 `overlay`，两个转换器都会转发给 `exportExcel`——三个入口定制或关闭遮罩的方式完全一致。见[表格与 ECharts → 与 exportExcel 的关系](/zh/packages/excel-exporter/api/05-table-and-echarts#与-exportexcel-的关系)。
- **给非 Excel 任务驱动同一遮罩。** 遮罩已内置于导出包，公开路径是 `@marcusok/excel-exporter/overlay` 子路径——导出通用的遮罩入口 `showExportOverlay`（与 Excel 无关）及句柄类型。文档示例见[进度遮罩 → 自己驱动遮罩](/zh/packages/excel-exporter/guide/11-overlay#自己驱动遮罩)。

## 版本、发布与兼容性

- **各包独立版本号，数字互不对应**——五个包由 Changesets 独立发版（`.changeset/config.json` 无 `fixed` / `linked` 分组）。不要从版本号读兼容性：exporter 3.x、preview 2.x，以及 word/PDF/PPTX 三个预览包的 0.x，互相之间什么也不说明。
- **它们出自同一个仓库**，同一套 CI、同一个 `main`。时间相近的发布构建自同一工作区状态，这才是真正的兼容性陈述。
- **全仓没有任何 `peerDependencies`**，没有 peer 版本区间需要你对齐。
- **引擎升级即业务包发版。** modern-xlsx 的升级落在内部引擎层，同时落在两个 Excel 业务包的构建期钉版 manifest 中（三处 manifest 一起动；运行时版本由引擎层唯一决定），随后以消费它的业务包的新版本发出——没有独立的引擎发布，也没有跨包版本联动。docx-preview 的升级只落在 word 预览包；pdfjs-dist 的升级只落在 PDF 预览包。

### 怎么判断两个版本搭配得上

与其看兼容矩阵，不如看清单——它是权威且始终最新的：

1. 查某包的依赖：`npm view @marcusok/excel-exporter dependencies`（或直接看 `node_modules` 里安装的 `package.json`）。
2. 版本间的行为变化读对应包的 `CHANGELOG.md`——由 Changesets 按包生成，每条记录恰好对应你要升级的那个范围。

## 我该装哪个包？

| 我需要……                                               | 安装                                                                                                |
| ------------------------------------------------------ | --------------------------------------------------------------------------------------------------- |
| 把数据导出为可下载的 `.xlsx`                           | [`@marcusok/excel-exporter`](/zh/packages/excel-exporter/)                                          |
| 从 Ant Design / Element Plus 表格或 ECharts 导出       | `@marcusok/excel-exporter`——适配器就在主入口                                                        |
| 只读展示上传的 `.xlsx` / `.xlsm` / `.csv`              | [`@marcusok/excel-preview`](/zh/packages/excel-preview/)                                            |
| 只读展示上传的 `.docx` / `.docm`                       | [`@marcusok/word-preview`](/zh/packages/word-preview/)                                              |
| 只读展示上传的 `.pdf`（懒加载滚动、文本可选中）        | [`@marcusok/pdf-preview`](/zh/packages/pdf-preview/)                                                |
| 只读展示上传的 `.pptx`（懒加载滚动、页导航、加密解密） | [`@marcusok/ppt-preview`](/zh/packages/ppt-preview/)                                                |
| 无头解析工作簿（SSR、自定义渲染器、框架封装）          | `@marcusok/excel-preview`——`parseWorkbookBytes`，无需 DOM                                           |
| 视觉保真预览 Word 文档（分页、脚注）                   | `@marcusok/word-preview`——注意它是主线程渲染，见其[限制](/zh/packages/word-preview/guide/02-limits) |
| 两者都要——预览文件并让用户导出修正版                   | 两个 Excel 包都装；各持一份引擎副本（见上文）                                                       |
| 给非 Excel 的任务用同一套进度遮罩                      | `@marcusok/excel-exporter`——`/overlay` 子路径导出通用遮罩入口                                       |
| 配置 WASM 托管（自托管 / CDN 副本）                    | 用哪个 Excel 包就配哪个——其导出的 `configureWasm` 覆盖该包的 loader 与 worker URL                   |
| 不碰引擎预览 `.csv`                                    | `@marcusok/excel-preview`——CSV 路径是纯 JS；`.xlsx` / `.xlsm` 解析始终走 WASM 引擎                  |

## 相关内容

- [生态](/zh/guide/)——大类与路线图。
- [快速上手](/zh/guide/01-getting-started)——导出包的安装命令与它带来的东西。
- [资产与自托管](/zh/packages/excel-preview/guide/02-assets)——WASM / worker 资产如何定位，以及自托管。
- [进度遮罩](/zh/packages/excel-exporter/guide/11-overlay)——`overlay` 选项与历史子路径。
