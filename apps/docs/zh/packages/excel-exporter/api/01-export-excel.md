# API：exportExcel 与配置

## exportExcel

```ts
exportExcel(options: ExportOptions): Promise<ExportResult>
```

核心入口函数（`exportTable` / `exportEcharts` 等便捷封装最终都委托给它）。根据数据量与环境自动路由到 main / worker / stream，WASM 不可用时降级无样式快速流。auto 模式阈值（浏览器）：< 20,000 行 → main · 20,000–49,999 → worker + Workbook · ≥ 50,000 → worker + 快速流；Node/SSR：< 50,000 → main · ≥ 50,000 → stream——见[自动模式路由](/zh/packages/excel-exporter/guide/03-auto-mode)。

## ExportOptions

| 字段         | 类型                                               | 必填 | 说明                                                                                                                                                                                                                                                                              |
| ------------ | -------------------------------------------------- | ---- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `sheets`     | `SheetConfig[]`                                    | 是   | 工作表配置，至少一个                                                                                                                                                                                                                                                              |
| `filename`   | `string`                                           | 是   | 下载文件名，不以 `.xlsx` 结尾时末尾自动追加（判定不区分大小写——`"REPORT.XLSX"` 不会再追加）                                                                                                                                                                                       |
| `mode`       | `"auto" \| "main" \| "worker" \| "stream"`         | —    | 默认 `"auto"`，按行数自动路由。该字段有枚举校验：传入其它值（如 `"Main"`）会以 `{ success: false }` 前置失败，而不是被静默当作 `auto`                                                                                                                                             |
| `onProgress` | `(progress: number) => void`                       | —    | 0 → 1；首尾 0 与 1 由 `exportExcel` 在所有路径各上报一次（含流式兜底与最终失败的导出）；分段进度仅 stream 路径有（每 1000 行上报一次）                                                                                                                                            |
| `onPhase`    | `(phase: ExportPhase, durationMs: number) => void` | —    | `init` / `build` / `download` 阶段耗时，在阶段**结束时**触发——上报规则见下文                                                                                                                                                                                                      |
| `download`   | `boolean`                                          | —    | 默认 `true` 触发浏览器下载；`false` 只返回 Blob                                                                                                                                                                                                                                   |
| `overlay`    | `boolean \| ProgressOverlayOptions`                | —    | 2.8.0 起**默认开启**。`false` 完全关闭全屏进度遮罩；传选项对象可定制——`delayMs`、`minVisibleMs`、`fadeOutMs`、`zIndex`、`container`、`blockInteraction`、`theme`、`text`（完整字段表见指南）。Node/SSR 下为空操作。见[进度遮罩指南](/zh/packages/excel-exporter/guide/11-overlay) |

阶段上报规则（`onPhase`）：

- **`init`**——WASM 初始化。主线程路径度量 `loader.ensureLoaded()`；worker 路径只在 worker 真正（重新）初始化时上报——已缓存的 WASM 实例不上报。主线程的无 WASM 流式路径上报零时长的 `init`，保持阶段序列稳定；浏览器的 worker-stream 路径（worker 内跑流引擎）则完全不上报 `init`。
- **`build`**——工作簿构建。每次**主线程**构建尝试各上报一个 `build`（在 `finally` 里，即使该次尝试抛错也上报——降级链中每次主线程尝试各有一个）。在 worker **内部**失败的尝试不上报 `build`；其失败只通过主线程重试的阶段与报错体现。
- **`download`**——同步的浏览器下载触发；仅当 `download !== false` **且**存在 `document` 时上报（Node 下永不上报）。

## ExportResult

| 字段        | 类型            | 说明                                                                                                                                                                                     |
| ----------- | --------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `success`   | `boolean`       | 是否成功                                                                                                                                                                                 |
| `blob?`     | `Blob`          | 导出文件内容                                                                                                                                                                             |
| `engine?`   | `"modern-xlsx"` | 实际使用的引擎                                                                                                                                                                           |
| `mode?`     | `ExportMode`    | 实际使用的模式                                                                                                                                                                           |
| `duration?` | `number`        | 完整导出耗时（ms）。主线程路径度量整次导出；worker 路径在主线程从调用 `exportInWorker` 起量（含 postMessage 前序列化与 worker 往返）直到 Blob 构建完成，因此**宽于** worker 内纯构建耗时 |
| `rowCount?` | `number`        | 导出的数据行数——不含表头行                                                                                                                                                               |
| `error?`    | `Error`         | `success` 为 `false` 时的失败原因。**同时**——与 `success: true` 一起——也会在经无样式快速流兜底**成功**的导出上设置，内容为降级原因；请先看 `success`：仅有 `error` 不代表导出失败        |

## configureWasm

```ts
configureWasm(options: LoaderOptions): void
```

可选——资产默认定位到随包发布的位置（见[安装与配置](/zh/packages/excel-exporter/guide/02-installation)）。仅自托管副本、CDN、或不支持资产 URL 的打包器需要覆盖。

| 字段              | 默认值             | 说明                                                                                                                                                  |
| ----------------- | ------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| `wasmUrl`         | 随包发布的 `.wasm` | 覆盖为自托管 / CDN 副本                                                                                                                               |
| `workerUrl`       | 随包发布的 worker  | 覆盖为**同源**自托管副本——跨域 URL 在 `Worker` 构造时即被浏览器拒绝，普通 CDN 在这里不可用                                                            |
| `timeoutMs`       | `10_000`           | 单次加载超时                                                                                                                                          |
| `maxRetries`      | `3`                | 最大加载尝试次数（默认共 3 次含首次，退避 300ms/600ms）                                                                                               |
| `workerTimeoutMs` | `120_000`          | Worker 导出超时。必须保持 > 0——没有「禁用」取值：`0`/负数意味着「立即超时」而非「不超时」                                                             |
| `parseWorkerUrl`  | —                  | 共享 `LoaderOptions` 类型上存在但对本包无效：它为 `@marcusok/excel-preview` 的解析 worker 指定自托管地址；本包的导出 worker 用上方的 `workerUrl` 配置 |

## 其他导出符号

以下成员均来自主入口，按"什么时候会用到"分组。

### 底层构建入口

- `WorkbookBuilder` —— `static create(): Promise<WorkbookBuilder>` · `addSheet(config: SheetConfig): this` · `toBuffer(): Promise<Uint8Array>` · `toBlob(): Promise<Blob>`。批量化构建、完整样式、不做模式路由。依赖 WASM：`create()` 内部会初始化引擎，无 WebAssembly 的环境中直接失败——与 `exportExcel` 不同，这个底层入口没有流式降级。**不展开** `SheetConfig.indexColumn`——直连时请传入已展开的 sheet（见下方 `applyIndexColumn`），否则该字段被静默忽略。
- `exportAsStream(sheets, onProgress?)` —— `(sheets: SheetConfig[], onProgress?: (p: number) => void) => Promise<{ bytes: Uint8Array; rowCount: number }>`。直接调用无样式快速流。第二个参数是**回调函数本身，不是选项对象**。`indexColumn` 的注意事项同 `WorkbookBuilder`。
- `applyIndexColumn(sheet)` / `INDEX_PROP` —— `(sheet: SheetConfig) => SheetConfig`，`INDEX_PROP = "__index__"`。把 `indexColumn` 展开为最左侧的保留列，即 `exportExcel` 内部所做的这一步；也是上面两个底层入口拿到同样结果唯一的公开手段。返回**新** sheet（不原地修改，`merges` 整体右移一列）；若用户列已占用该保留 prop 则抛错。
- `exportTable(options)` —— `(options: TableExportOptions) => Promise<ExportResult>`。表格形态数据的便捷入口。完整参考：[表格与 ECharts 入口](./05-table-and-echarts)。
- `exportEcharts(options)` —— `(options: EChartsExportOptions) => Promise<ExportResult>`。ECharts 数据的便捷入口。同一参考页。

### 样式与格式化工具

| 导出                                                           | 签名                                                     | 说明                                                                                                                                                                            |
| -------------------------------------------------------------- | -------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `StylePresets` / `StylePresetName`                             | `StylePresets` 对象；类型名是它的键联合                  | 共 8 个预设：`header` `currency` `percent` `date` `datetime` `dataRow` `bordered` `danger`                                                                                      |
| `BaseCellStyle`                                                | `CellStyle`（已冻结）                                    | 库级基底样式，所有单元格样式都铺在它之上（见[样式指南](/zh/packages/excel-exporter/guide/05-styles)）                                                                           |
| `applyFormat(value, spec)`                                     | `(value: unknown, spec: FormatSpec) => string \| number` | 对单个值应用 `FormatSpec`                                                                                                                                                       |
| `toStr(value)`                                                 | `(value: unknown) => string`                             | 导出器自己用于未格式化单元格的值字符串化：对象 → JSON（循环引用与怪异 `toJSON` 兜底为 `String()`）、`null`/`undefined` → `""`、Date → ISO、bigint/布尔/symbol/函数 → `String()` |
| `numFormatForSpec(spec)`                                       | `(spec: FormatSpec) => string \| null`                   | 某个 `FormatSpec` 解析出的 Excel 数字格式代码                                                                                                                                   |
| `formatDateByPattern(value, pattern)`                          | `(value: unknown, pattern: string) => string`            | 按显式 pattern 格式化日期 / 日期时间                                                                                                                                            |
| `displayValue(col, row)` / `resolveCellFormat(col, row)`       | `(col: ColumnConfig, row: Record<string, unknown>) => …` | 导出器写进单元格的值、以及解析出的格式——可复用于自研预览或导出前预检                                                                                                            |
| `validateSheetName(name)` / `validateMerges(sheet, leafCount)` | 返回 `void`，非法输入抛错                                | `exportExcel` 自身执行的前置校验；可在不导出的情况下预检输入                                                                                                                    |
| `DEFAULT_DATE_PATTERN` / `DEFAULT_DATETIME_PATTERN`            | `"yyyy-MM-dd"` / `"yyyy-MM-dd HH:mm"`                    | 未显式给 `pattern` 时 `{ type: "date" }` / `{ type: "datetime" }` 的回退格式                                                                                                    |

### WASM 加载器

`getWasmLoader()` 返回本包的 `WasmLoader` 单例。观察它请用 `isReady` / `supported` 两个 getter 与 `getOptions()`——`LoadState` 状态机（`idle` / `loading` / `ready` / `error`）是私有状态，不是可读字段。它与上文的 `configureWasm` 都来自打包进本包 `dist` 的内部引擎层（`@marcusok/excel-preview` 也为它自己打包同一层——两包各持一个 loader 实例）。

### 遮罩

`ProgressOverlayOptions` / `ProgressOverlayTextOptions` / `ProgressOverlayHandle` 被重导出，便于在不引入第二个 import 的前提下给 `overlay` 选项标注类型。遮罩本体已内置于本包；见[进度遮罩指南](/zh/packages/excel-exporter/guide/11-overlay)。

### 子路径

| 子路径                                                                       | 提供什么                                                                                                                                                                                                                                                                                                                                                                                                                             |
| ---------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `@marcusok/excel-exporter/styles`                                            | 独立的 `StylePresets` + `StylePresetName`——取值与主入口一致，供不能引入整个导出包的场景按需取用                                                                                                                                                                                                                                                                                                                                      |
| `@marcusok/excel-exporter/worker-utils`                                      | `exportInWorker(options: ExportOptions, mode: "workbook" \| "stream"): Promise<ExportResult>` 与 `terminateWorker(): void`。**主入口不提供**，且 `mode` 为必填。`terminateWorker()` 会终止共享 worker，并以 `worker terminated` 拒绝所有在途请求                                                                                                                                                                                     |
| `@marcusok/excel-exporter/overlay`                                           | 2.8 之前的历史兼容层 + 通用遮罩入口：`exportExcelWithOverlay(options, overlay?)`、`showExportOverlay`（即 `showProgressOverlay` 的别名——progress-overlay 层已内联进导出包、不再单独发包）、`nextPaint()`（该工具唯一的公开入口——主入口不导出它）以及重命名类型 `ExportOverlayHandle` / `OverlayOptions` / `OverlayTextOptions`。注意优先级——封装的 `overlay` 参数默认值为 `{}`，因此它会**覆盖** `options` 里已有的 `overlay: false` |
| `@marcusok/excel-exporter/dist/export.worker.js` · `…/dist/modern-xlsx.wasm` | 原始 worker / WASM 资产，供不想走加载器默认 URL 解析的自托管场景使用                                                                                                                                                                                                                                                                                                                                                                 |

```ts
import {
  exportExcel,
  configureWasm,
  WorkbookBuilder,
  exportAsStream,
  exportTable,
  exportEcharts,
  getWasmLoader,
} from "@marcusok/excel-exporter";
```

> 本文档覆盖常用的稳定 API。主入口另外还重导出了上表列出的 `format-utils` 工具、`BorderStyle`（声明于本包自己的 `types.ts`）以及引擎层类型 `LoaderOptions` / `LoadState`（可从主入口以类型方式 import）；完整导出清单见 [GitHub 上的 `src/index.ts`](https://github.com/yourbusiness/marcusok/blob/main/packages/excel-exporter/src/index.ts)（安装包内对应 `dist/index.d.ts`）。
