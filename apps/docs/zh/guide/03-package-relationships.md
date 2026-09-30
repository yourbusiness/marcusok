# 包之间的关系与选型

生态目前对外发布四个包，分属三大类（分类定义见[生态介绍](/zh/guide/)）。本文回答选型时最常见的几类问题：各包分别负责什么、谁依赖谁、导出与预览能否同页共存，以及版本之间是什么关系。

## 四个发布包

| 包                                                             | 分类 | 职责                                                                 | 运行时依赖                                          |
| -------------------------------------------------------------- | ---- | -------------------------------------------------------------------- | --------------------------------------------------- |
| [`@marcusok/excel-exporter`](/zh/packages/excel-exporter/)     | 导出 | 业务数据 → `.xlsx` 下载：样式、格式化、worker/流式路由与多级降级兜底 | `@marcusok/xlsx-core`、`@marcusok/progress-overlay` |
| [`@marcusok/excel-preview`](/zh/packages/excel-preview/)       | 预览 | 只读渲染 `.xlsx` / `.xlsm` / `.csv`：worker 解析 + 虚拟滚动          | `@marcusok/xlsx-core`                               |
| [`@marcusok/xlsx-core`](/zh/packages/xlsx-core/)               | 共享 | 全仓唯一的 modern-xlsx 引擎层：WASM 加载、资产分发、稳定再导出面     | `modern-xlsx`（构建期打包进自身 `dist`）            |
| [`@marcusok/progress-overlay`](/zh/packages/progress-overlay/) | 共享 | 与框架无关的全屏进度遮罩（旋转圆环 / 百分比条、双主题）              | 无                                                  |

依赖关系很小，且是刻意设计的：

```
@marcusok/excel-exporter ──┬──> @marcusok/xlsx-core
                            └──> @marcusok/progress-overlay

@marcusok/excel-preview  ───────> @marcusok/xlsx-core
```

由此有三条结论值得记住：

- **导出与预览是平级关系，不是上下层。** 两者互不依赖，任何一个都能单独安装；导出引擎的实现细节不会渗进预览 API，预览也不是为导出包服务的附属品。
- **`@marcusok/xlsx-core` 是唯一的共享引擎。** 两个业务包都直接依赖它——这正是两者引擎完全一致的原因（见下文"共享同一个引擎"一节）。
- **`@marcusok/progress-overlay` 是叶子节点。** 它自身零依赖，完全不知道 Excel 的存在，且只被导出包使用——预览不会把它带进依赖树。它是一个通用 UI 包，只是恰好被导出包的 `overlay` 选项驱动（[进度遮罩](/zh/packages/excel-exporter/guide/11-overlay)）。

### 各自什么时候用

- **`@marcusok/excel-exporter`** —— 数据已经在应用里（表格、数组、ECharts 配置），需要产出一个供用户下载的文件。从[快速开始](/zh/guide/01-getting-started)入手；表格与 ECharts 的便捷适配器见[表格与 ECharts 便捷入口](/zh/packages/excel-exporter/api/05-table-and-echarts)。
- **`@marcusok/excel-preview`** —— 用户上传了 `.xlsx`，需要**展示**它，且不需要编辑。从[快速开始](/zh/packages/excel-preview/guide/01-quick-start)入手。
- **`@marcusok/xlsx-core`** —— 通常作为依赖被动装上。只有当你要在 modern-xlsx 的接口之上自建集成，或者想在不引入任何文档包的前提下配置 WASM loader 时才需要直接引用它（loader 就在这个包里，见 [configureWasm](/zh/packages/excel-exporter/api/01-export-excel#configurewasm)）。
- **`@marcusok/progress-overlay`** —— 你想把同一套遮罩用在与 Excel 无关的任务上。用法见[包指南](/zh/packages/progress-overlay/guide/01-usage)。

## 导出与预览能同页共存吗

能——两者本就是为共存设计的，它们之间唯一的交集就是共享引擎。同时安装后，主线程上只会得到**一份引擎实例与一份 WASM 二进制**，而不是两份：两个包的主入口构建都把 `@marcusok/xlsx-core` 保持为 external（不各自打包一份），因此运行时解析到同一个模块。这一点在预览侧有明确说明，见[安装](/zh/packages/excel-preview/#安装)与[资产与自托管 → 默认定位机制](/zh/packages/excel-preview/guide/02-assets#默认定位机制)。

### 哪些是共享的，哪些不是

| 项目                      | 共享还是各包一份                                                                                        |
| ------------------------- | ------------------------------------------------------------------------------------------------------- |
| 主线程引擎与 WASM         | **共享**——主线程只有一份 `@marcusok/xlsx-core`、一份 WASM 二进制                                        |
| WASM loader 与配置        | **共享**——`configureWasm` / `getWasmLoader` 在两个包里是同一批对象                                      |
| Worker 脚本               | **各包一份**——`export.worker.js`（导出）与 `parse.worker.js`（预览）是两个不同的文件，各自随包发布      |
| Worker URL 配置字段       | **各包独立字段**——`workerUrl` 属于导出 worker，`parseWorkerUrl` 属于解析 worker。单字段无法同时服务两者 |
| Worker **内部**的引擎副本 | **各包一份**——无法避免，原因见下                                                                        |
| 遮罩 UI                   | **仅导出包**——预览不依赖遮罩包                                                                          |

最容易被忽略的是 worker 这一项。浏览器的 module worker 无法解析裸导入（import map 在 `WorkerGlobalScope` 中不生效），因此每个 worker 脚本必须是自包含的单文件——也就意味着**每个 worker 各自内置一份引擎**。这份重复来自 worker 边界本身，不是设计疏漏；"共享单份引擎"是对主线程的保证。预览的资产页对此有明确说明，构建 worker 的代码也正因如此拒绝共用同一个 URL 字段（[worker-client.ts](https://github.com/yourbusiness/marcusok/blob/main/packages/excel-preview/src/worker/worker-client.ts) 读 `parseWorkerUrl`，`worker-exporter.ts` 读 `workerUrl`）。

所以在同时使用两个包的页面上，最多出现：一份主线程引擎、一个导出 worker、一个解析 worker。实际业务中一个页面很少会同时跑两条 worker 路径。

### 实际影响

- **一次 `configureWasm` 同时配置两者。** 全站只有一个 loader，自托管场景下一次调用同时传入 `wasmUrl` 与两个 worker URL 即可——预览的[自托管 / CDN](/zh/packages/excel-preview/guide/02-assets)一节正是为这种场景写的。
- **超时设置同样是共享的。** `timeoutMs`、`maxRetries`、`workerTimeoutMs` 都挂在同一个 loader 上，谁加载资产、谁派发 worker，就由谁生效。
- **降级策略各包独立。** 导出包有自己的无样式快速流式兜底，预览会回退到主线程解析；一个包进入兜底不会改变另一个包的路由。
- **不需要为了解析而装导出包，也不需要为了写入而装预览包。** 共享引擎不等于共享 API。

## 共享同一个引擎（`xlsx-core`）

引擎包自身有[独立文档页](/zh/packages/xlsx-core/)；下面从"选型"的角度说版本与依赖层面的含义。

### 它是普通依赖，不是 peerDependency

两个业务包都把引擎声明在普通的 `dependencies` 里：

- `packages/excel-exporter/package.json` —— `"@marcusok/xlsx-core": "workspace:*"`（同时还有 `@marcusok/progress-overlay`）
- `packages/excel-preview/package.json` —— `"@marcusok/xlsx-core": "workspace:*"`（它唯一的运行时依赖）

全仓没有任何 `peerDependencies` 声明。`workspace:*` 是开发期协议，发布时会被改写为具体版本（目前两个包在已发布的清单里都把 `@marcusok/xlsx-core` 钉到确切版本，可用 `npm view @marcusok/excel-exporter dependencies` 自行核对）。

这个选择对使用方的直接意义是：**你永远不必操心该装哪个引擎版本。**`pnpm add @marcusok/excel-preview` 带进来的就是它构建与测试时所用的引擎；升级引擎等于升级你已经在用的那个包，而不是往自己的依赖清单里再加一项，也就不会与导出包所用的引擎发生范围冲突。

### 为什么"再导出面"能保证单一版本来源

modern-xlsx 的 JS 胶水与 `fflate` 在构建期被打包进 `xlsx-core` 的 `dist`，而 `xlsx-core` 本身只是 loader 加一层对 modern-xlsx 的稳定再导出面。业务包只通过这层再导出面使用引擎——导出包甚至把 `configureWasm` / `getWasmLoader` 也再导出出去，以保证自己的公开 API 不变。结果是：

- 全仓只有一处声明 modern-xlsx 的版本。
- 两个包看到的是同一份引擎构建、同一套错误类型、同一个 WASM 二进制默认 URL。

### `getWasmLoader()` 就是同一个对象

`xlsx-core` 暴露一个模块级的 loader 实例，由 `getWasmLoader()` 返回；两个包都再导出了这个函数，所以无论从 `@marcusok/excel-exporter` 还是从 `@marcusok/excel-preview` 引入 `configureWasm` / `getWasmLoader`，配置与观察到的都是**同一个 loader**——预览 API 参考里也写明了这点（见 [createPreview → 其他导出](/zh/packages/excel-preview/api/01-create-preview#其他导出)）。

> "单实例"是模块实例层面的保证：只要打包器解析到的是同一份 `@marcusok/xlsx-core` 就成立。上面的确切版本钉法让这成为常态；如果同一个产物里出现了两份该包，那就是两份引擎、两套 loader 状态。

### `xlsx-core` 的构建期身份

在本仓库内，`xlsx-core` 同时也是**构建期**依赖：它发布了一条 `@marcusok/xlsx-core/tsup` 子路径，内含两个业务包 `tsup.config.ts` 共用的 esbuild 插件，使两者在打包引擎时不会各自漂移。应用侧不会引用这条子路径。

此外，两个业务包都会在构建后把 WASM 二进制转发一份到自己的 `dist`——这是对旧深路径导入的兼容转发，规范位置是核心包的 `dist/modern-xlsx.wasm`，也是 loader 的默认指向。

## `progress-overlay` 与导出包的关系

遮罩是唯一一处"共享包被接进业务包默认行为"的地方。

- 自 **2.8.0** 起，每次导出默认显示遮罩：无需额外 import，也无需包装函数。
- `overlay` 字段声明在 `ExportOptions` 上，有三态——省略/`true`（使用导出默认文案）、`false`（完全关闭：不挂载任何节点、不多让一帧）、传 `ProgressOverlayOptions` 对象（自定义文案、`delayMs`、主题等）。三态取值与自定义文案的合并语义详见[进度遮罩 → 取值](/zh/packages/excel-exporter/guide/11-overlay#取值)，本文不重复。
- **该字段只存在于 `ExportOptions` 上。** `TableExportOptions` 与 `EChartsExportOptions` 都没有声明它，两者的转换器也不转发它，因此这两个便捷入口既不能定制、也不能关闭遮罩：它们始终按默认（开启）运行，要显式控制只能走主入口 `exportExcel`（细节见[表格与 ECharts 便捷入口](/zh/packages/excel-exporter/api/05-table-and-echarts)）。
- **2.8 之前的子路径仍然保留。** `@marcusok/excel-exporter/overlay` 现在只是一层薄封装（`exportExcelWithOverlay`），等价于 `exportExcel({ ...options, overlay })`；注意能力搬进共享包时文案结构与句柄方法名都变了，迁移说明见[进度遮罩 → 历史子路径](/zh/packages/excel-exporter/guide/11-overlay#历史子路径)。
- **该包可以完全独立使用。** 如果你自己的流程需要同一套遮罩，直接引入 `@marcusok/progress-overlay` 驱动句柄即可，文档中的示例见[进度遮罩 → 自己驱动遮罩](/zh/packages/excel-exporter/guide/11-overlay#自己驱动遮罩)；这个包与 Excel 没有任何耦合。

## 版本、发布与兼容性

- **各包独立版本号，且数字并不对齐**——四个包由 Changesets 独立发版（`.changeset/config.json` 未声明 `fixed` / `linked` 分组）。不要从版本号推断兼容性：导出包 2.x、预览 1.x、核心 1.x、遮罩 1.x 之间没有对应关系。
- **它们出自同一个仓库**，跑同一套 CI、同在 `main` 分支发布。发布时间接近的版本来自同一份工作区状态——这才是真正的兼容性依据；这几个包不是碰巧同名 scope 的独立项目。
- **全仓没有 peerDependencies**，因此不存在需要你调和的 peer 范围；业务包把它构建时所用的引擎版本直接钉在自己的依赖里。
- **引擎层升级会自动向上传导。** `xlsx-core` 发版时，依赖方的依赖项会随之前移并发版——配置项 `updateInternalDependencies: "patch"`（`.changeset/config.json`）保证了这点，于是引擎发版是以"导出包/预览包的新版本"的形式到达使用方，而不是要求使用方自己动依赖。仓库内的构建顺序也由依赖关系驱动（Turborepo 的 `^build`），业务包不会基于过期的核心包构建。

### 怎么判断两个版本搭配得上

与其列一张兼容矩阵，不如直接看清单文件，它永远权威且最新：

1. 看某个包钉的是什么：`npm view @marcusok/excel-exporter dependencies`（或直接读 `node_modules` 下已安装的 `package.json`）。
2. 确认它钉的引擎版本就是实际安装的版本——用 `pnpm add` 正常安装文档包时没有任何需要调和的地方；只有当你在自己的清单里手工钉了引擎版本时才会冲突。
3. 想看两个版本之间的行为变化，读该包的 `CHANGELOG.md`：它由 Changesets 按包生成，每条记录都精确对应你正在升级的那段范围。

## 我该装哪个包？

| 我要做的事                                               | 装什么                                                                               |
| -------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| 把数据导出为可下载的 `.xlsx`                             | [`@marcusok/excel-exporter`](/zh/packages/excel-exporter/)                           |
| 从 Ant Design / Element Plus 表格或 ECharts 配置导出     | `@marcusok/excel-exporter`——适配器已随主入口发布                                     |
| 只读展示用户上传的 `.xlsx` / `.xlsm` / `.csv`            | [`@marcusok/excel-preview`](/zh/packages/excel-preview/)                             |
| 无界面解析工作簿（SSR、自建渲染器、框架封装）            | `@marcusok/excel-preview` 的 `parseWorkbookBytes`，不需要 DOM                        |
| 两者都要：既能预览文件，又让用户导出订正后的副本         | 两个文档包都装；引擎与 WASM 二进制只会加载一次                                       |
| 给非 Excel 的任务用同一套进度遮罩                        | [`@marcusok/progress-overlay`](/zh/packages/progress-overlay/)                       |
| 统一配置 WASM 托管地址（无论最终用了哪几个包）           | 任选一个文档包（导出的 `configureWasm` 是共享的），或直接用 `@marcusok/xlsx-core`    |
| 直接在 modern-xlsx 之上自建（自己的 builder / 读取链路） | [`@marcusok/xlsx-core`](/zh/packages/xlsx-core/)                                     |
| 完全不碰引擎就预览 `.csv`                                | `@marcusok/excel-preview`——CSV 路径是纯 JS；`.xlsx` / `.xlsm` 解析始终使用 WASM 引擎 |

## 相关内容

- [生态介绍](/zh/guide/)——分类与路线图。
- [快速开始](/zh/guide/01-getting-started)——导出包的安装与它会带进来的东西。
- [资产与自托管](/zh/packages/excel-preview/guide/02-assets)——两个包共用的 loader 及其资产定位机制。
- [进度遮罩](/zh/packages/excel-exporter/guide/11-overlay)——`overlay` 选项与历史子路径。
