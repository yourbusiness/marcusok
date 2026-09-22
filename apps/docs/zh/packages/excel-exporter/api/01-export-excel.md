# API：exportExcel 与配置

## exportExcel

```ts
exportExcel(options: ExportOptions): Promise<ExportResult>
```

核心入口函数（`exportTable` / `exportEcharts` 等便捷封装最终都委托给它）。根据数据量与环境自动路由到 main / worker / stream，WASM 不可用时降级无样式快速流。

## ExportOptions

| 字段         | 类型                                               | 必填 | 说明                                                                                                                                   |
| ------------ | -------------------------------------------------- | ---- | -------------------------------------------------------------------------------------------------------------------------------------- |
| `sheets`     | `SheetConfig[]`                                    | 是   | 工作表配置，至少一个                                                                                                                   |
| `filename`   | `string`                                           | 是   | 下载文件名，不以 `.xlsx` 结尾时末尾自动追加                                                                                            |
| `mode`       | `"auto" \| "main" \| "worker" \| "stream"`         | —    | 默认 `"auto"`，按行数自动路由                                                                                                          |
| `onProgress` | `(progress: number) => void`                       | —    | 0 → 1；首尾 0 与 1 由 `exportExcel` 在所有路径各上报一次（含流式兜底与最终失败的导出）；分段进度仅 stream 路径有（每 1000 行上报一次） |
| `onPhase`    | `(phase: ExportPhase, durationMs: number) => void` | —    | `init` / `build` / `download` 阶段耗时                                                                                                 |
| `download`   | `boolean`                                          | —    | 默认 `true` 触发浏览器下载；`false` 只返回 Blob                                                                                        |

## ExportResult

| 字段        | 类型            | 说明                         |
| ----------- | --------------- | ---------------------------- |
| `success`   | `boolean`       | 是否成功                     |
| `blob?`     | `Blob`          | 导出文件内容                 |
| `engine?`   | `"modern-xlsx"` | 实际使用的引擎               |
| `mode?`     | `ExportMode`    | 实际使用的模式               |
| `duration?` | `number`        | 完整导出耗时（ms）           |
| `rowCount?` | `number`        | 导出行数                     |
| `error?`    | `Error`         | 失败原因（兜底路径也会返回） |

## configureWasm

```ts
configureWasm(options: LoaderOptions): void
```

可选——资产默认定位到随包发布的位置（见[安装与配置](/zh/packages/excel-exporter/guide/02-installation)）。仅自托管副本、CDN、或不支持资产 URL 的打包器需要覆盖。

| 字段              | 默认值             | 说明                                                    |
| ----------------- | ------------------ | ------------------------------------------------------- |
| `wasmUrl`         | 随包发布的 `.wasm` | 覆盖为自托管 / CDN 副本                                 |
| `workerUrl`       | 随包发布的 worker  | 覆盖为自托管 / CDN 副本                                 |
| `timeoutMs`       | `10_000`           | 单次加载超时                                            |
| `maxRetries`      | `3`                | 最大加载尝试次数（默认共 3 次含首次，退避 300ms/600ms） |
| `workerTimeoutMs` | `120_000`          | Worker 导出超时                                         |

## 其他导出符号

- `WorkbookBuilder.create()` + `addSheet(config)` + `toBuffer()` / `toBlob()`：批量化构建，完整样式。**不展开** `SheetConfig.indexColumn`——直连时请传入已展开的 sheet（`applyIndexColumn(sheet)`），见 `IndexColumnOptions`；
- `exportAsStream(sheets, onProgress?)`：底层流式导出，返回 `Promise<{ bytes, rowCount }>`。`indexColumn` 的注意事项同 `WorkbookBuilder`；
- `applyIndexColumn(sheet)` / `INDEX_PROP`：把 `indexColumn` 展开为最左侧的保留 `__index__` 列（`exportExcel` 内部所做的正是这一步，也是上面两个底层入口唯一的公开补偿手段）；
- `exportTable(options)`：常见表格数据便捷导出，支持 Element Plus `prop`/`label`（即本库命名）、AntD `dataIndex`/`title` 与旧名 `key`/`header`；sheet 名默认 `"Sheet1"`（可用 `sheetName` 覆盖）；`freezeRows` / `autoFilter` / `merges` / `dataStyle` / `indexColumn` 会透传给 sheet；
- `exportEcharts(options)`：常见 ECharts 数据便捷导出，支持类目轴多系列（类目取自 `xAxis.data`；水平条形图取 `yAxis.data`）、饼图 `name/value`、散点数据两种写法（`[x,y]`、`[x,y,...dims]`——多维散点取前两维作 `X`/`Y`，被丢弃的额外维度以一次 console.warn 告知，或 `{ value: [x,y] }`）。`dataset` 模式与多根 x/y 轴会被明确报错拒绝。`layout` 可选 `"wide"`（默认，每系列一列）或 `"long"`（每系列-类目对一行），**仅对类目轴布局有意义**——item 数据（饼图/散点）会忽略它；`xAxis.data` 为空或缺失时走 item（名称/数值）布局。默认 sheet 名（`图表数据`）与表头为中文，可通过 `sheetName` / `seriesHeader` / `categoryHeader` / `nameHeader` / `valueHeader` 覆盖；散点布局的坐标表头是字面量 `X` / `Y`。long/item 布局下表头兼作行键，重复表头会被明确拒绝；
- `getWasmLoader()`：访问全局 WASM 加载器（状态：idle / loading / ready / error）。

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

> 入口还重导出了若干底层工具与类型（如 `format-utils` 的 `applyFormat` / `validateSheetName`、`LoaderOptions` / `LoadState`、`BorderStyle` 等），本文档只覆盖常用的稳定 API，完整列表见 `src/index.ts`。另有三条用于按需拆分的子路径：`@marcusok/excel-exporter/styles`（独立的 `StylePresets` 入口，主入口也重导出它）、`@marcusok/excel-exporter/worker-utils`（`exportInWorker` / `terminateWorker`，源入口 `src/worker-exporter.ts`——主入口不提供）与 `@marcusok/excel-exporter/overlay`（`exportExcelWithOverlay` / `showExportOverlay`，源入口 `src/overlay.ts`——见[进度遮罩指南](/zh/packages/excel-exporter/guide/11-overlay)）。
