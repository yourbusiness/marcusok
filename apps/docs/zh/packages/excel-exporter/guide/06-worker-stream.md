# Worker 与流式

## Worker 多线程（≥ 20,000 行或 ≥ 250,000 格）

浏览器中数据量 ≥ 20,000 行**或 ≥ 250,000 估算单元格**时，`auto` 会选择 Worker 路径：主线程只做一次结构化克隆，其余工作都在 Worker 内执行。克隆成本随单元格数增长（10 万×6 约 94ms、40k×50 约 200ms，Chromium 实测）。低于 stream 门槛（行 < 50,000 **且** 格 < 1,500,000）时，Worker 内加载 WASM 并走 Workbook 路径；行 ≥ 50,000 **或** 格 ≥ 1,500,000 时切换为不依赖 WASM 的 Fast stream（见下）。

worker 资产（`export.worker.js`，自包含单文件 ESM）默认自动定位，无需任何配置；仅自托管副本场景需要 `configureWasm({ workerUrl })` 覆盖。

Worker 路径行为：

- **Worker 失败会优雅降级**——Worker 路由失败时（例如 `workerUrl` 覆盖配置指向 404），导出会**先回退到主线程重试**（modern-xlsx 保留样式；达到 stream 门槛的 Fast stream 本身不依赖 WASM）。在 Worker + Workbook 路由上，重试再失败才继续降级到无样式的流式兜底（`mode: "stream"` 且 `result.error` 非空，`success` 仍为 `true`）；而在 Worker + stream 路由上（浏览器达到 stream 门槛），主线程重试本身就是对同一份输入跑同一个 fast stream，因此再失败即终局——导出以 `success: false` 结束，不再尝试第三次构建。无论哪条路由，调用方的 Promise 都正常 resolve（不会 reject），每一级降级都会在 console 打印 `[excel-exporter]` 前缀警告；
- Worker 实例复用，但请求经**主线程 FIFO 队列串行化**：同一时刻只有一个导出在处理。每个请求的超时从派发时开始计时（排队时间不计），超时/中止只收尾当前请求——排队中的导出落到新 worker 继续执行；
- **函数形式的 format 会被剥离**（结构化克隆无法传递函数）——worker 路径请使用 FormatSpec；
- `onProgress` / `onPhase` 会从 Worker 转发回主线程：`build` 在 Worker 内构建完成时上报；`init` 仅在 Worker + Workbook 路由、且 Worker 确实执行了 WASM 初始化时才上报——Worker + stream 路由不使用 WASM，完全不上报 `init`。

## 流式写入（≥ 50,000 行或 ≥ 1,500,000 格）

`fast-xlsx.ts` 使用 `fflate` 生成 minimal OOXML，10 万行约 0.8s（对比 Workbook 路径 Chrome 实测 17.5s——同一断崖在 Node 冷启动约 21s，见[性能页](/zh/packages/excel-exporter/guide/07-performance)）。`auto` 在行 ≥ 5 万**或**格 ≥ 150 万时自动选择它。

Stream 路径的已知限制：

| 特性                      | Stream 路径                                                              |
| ------------------------- | ------------------------------------------------------------------------ |
| 多行表头（`children`）    | 支持（表头自动合并）                                                     |
| 单元格合并（`merges`）    | 支持（数据区）                                                           |
| 单元格样式（`style`）     | 不支持                                                                   |
| 表头样式（`headerStyle`） | 不支持                                                                   |
| 列宽（`width`）           | 不支持                                                                   |
| 冻结行 / 自动筛选         | 不支持                                                                   |
| FormatSpec 数字/日期格式  | 支持（最小 `styles.xml` / 仅 numFmt；`CellStyle.numFormat` 会被丢弃）    |
| 日期格式                  | 写序列号 + numFmt（格式码与 Workbook 一致，除非 `style.numFormat` 冲突） |
| 进度回调                  | 每 1000 行上报一次                                                       |

被跳过的特性（单元格样式、表头样式、列宽、冻结等）会在 console 打印 `[excel-exporter] stream mode: features not supported (...)` 警告，并随 `result.degradedFeatures` 返回——正常的 auto 大文件路由同样会设置该字段（此时 `success` 仍为 `true`、`error` 为空）。失败驱动的降级还会同时设置 `error` 说明原因。

## 什么时候该用哪个

| 需求                                                 | 推荐路径                           |
| ---------------------------------------------------- | ---------------------------------- |
| 低于两个门槛（行 < 5 万且格 < 150 万）且需要完整样式 | `auto`（main / worker + Workbook） |
| 行 ≥ 5 万或格 ≥ 150 万，样式可接受降级               | `auto`（worker + Fast stream）     |
| Node 服务端大批量                                    | `auto`（main → 达门槛后 stream）   |
| 对主线程零阻塞有强要求                               | 显式 `mode: "worker"`              |

## 直接使用底层 API

库同时导出 `WorkbookBuilder` 与 `exportAsStream`，可在复杂场景下精细控制：

```ts
import { WorkbookBuilder, exportAsStream } from "@marcusok/excel-exporter";

// 批量化构建
const builder = await WorkbookBuilder.create();
builder.addSheet(sheetA).addSheet(sheetB);
const bytes = await builder.toBuffer();

// 流式导出
const { bytes, rowCount } = await exportAsStream(sheets, onProgress);
```

::: warning
这两个底层入口只识别**已展开**的 `__index__` 列——与 `exportExcel` 不同，
它们不会自行展开 `SheetConfig.indexColumn`，直接传入时该配置会被静默忽略。
若在这些路径上需要序号列，须先调用 `applyIndexColumn(sheet)`（即
`exportExcel` 内部使用的同一工具）。详见
[API 参考中的 `applyIndexColumn`](/zh/packages/excel-exporter/api/01-export-excel)。
:::
