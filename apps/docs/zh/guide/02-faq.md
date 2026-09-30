# 常见问题

这里按包汇总最容易先撞上的问题；每个包的文档里写得更细，入口都在[生态介绍](/zh/guide/)页。

## 通用

### 数据会被上传吗？

不会。所有处理都在浏览器或 Node 进程内完成，业务数据不离开本机。这一点对每个包都成立——导出、预览与引擎层都一样。

## Excel 导出（`@marcusok/excel-exporter`）

### 浏览器报 WASM 404

默认（零配置）定位下，Vite / webpack 5 会把随包发布的 `modern-xlsx.wasm` 自动发射为 hash 资产，不应出现 404。若仍遇到，通常是覆盖了 URL（`configureWasm({ wasmUrl })` 指向了错误路径），或使用的打包器不支持 `new URL(资产, import.meta.url)` 资产模式；把 `configureWasm` 指向站点实际可访问的地址，或从本包 `dist/` 把文件拷贝到静态目录即可。

### 构建产物正常，只有 `vite dev` 下 WASM 加载失败

Vite 开发服务器会把依赖预打包进 `node_modules/.vite/deps/`，在那里 `import.meta.url` 已不再指向包的 `dist/`，默认 WASM 地址因此指向不存在的路径。这个失败很安静——导出照样成功，但样式、列宽、冻结与筛选全部被剥离（降级到无样式快速流）。把包排除出预打包并重启 dev server：

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ["@marcusok/excel-exporter"] },
});
```

`vite build` 不受影响。完整诊断（含 `expected magic word` 编译报错，以及后续重试报 `WASM load previously failed` 的诊断陷阱）见[预打包注意事项](/zh/packages/excel-exporter/guide/02-installation#vite-开发服务器-预构建注意事项)。

### Worker 模式回退到了主线程

worker 资产（`export.worker.js`）默认自动定位；回退发生在 Worker 路由失败时（例如 `workerUrl` 覆盖配置指向了 404 的地址）。此时导出会**在主线程重试**（modern-xlsx 保留样式；Fast stream 本身不需要 WASM）——只有主线程重试也失败时，才最后降级到无样式的流式兜底。查看 console 中 `[excel-exporter]` 前缀的警告可定位原因。

### 导出成功但 result.error 有值

说明导出降级到了无样式快速流（WASM 失败或环境不支持）：样式被剥离，表头与合并保留。原因见 `result.error.message` 与 console 中的 `[excel-exporter]` 警告——通常是 wasm 资产 404。详见 [兜底机制](/zh/packages/excel-exporter/guide/08-fallback)。

### 10 万行数据导出非常慢（>15s）

大概率走了 `main` + `Workbook.toBuffer()` 路径——该路径在 ~5.5 万行后出现性能断崖。把 `mode` 保持为 `auto`（10 万行约 0.8s），或显式指定 `mode: "stream"` / `mode: "worker"`。详见 [自动模式路由](/zh/packages/excel-exporter/guide/03-auto-mode)。

### Stream 模式下样式不生效

Stream 路径支持多行表头（`children`）与数据区合并（`merges`），但不支持单元格样式、表头样式与列宽/冻结/筛选等布局特性（会在 console 打印警告）。需要完整样式时，控制在 5 万行以内走 Workbook 路径。详见 [Worker 与流式](/zh/packages/excel-exporter/guide/06-worker-stream)。

### 日期列显示为长文本，Excel 不识别为日期

不声明 `format` 时，`Date` 值会按普通文本写入单元格，Excel 不会识别为日期：所有路径（main / worker / stream）统一写入 ISO 字符串（如 `2026-07-01T00:00:00.000Z`）。日期列需要声明 `format: { type: "date" }`（或 `datetime`）：Workbook 路径会写入 Excel 日期序列并自动注入对应 `numFormat`，单元格才会被 Excel 识别为真正的日期。

## Excel 预览（`@marcusok/excel-preview`）

### 加密文件报 `PASSWORD_PROTECTED`

说明该文件已加密（Agile AES-256）且未提供口令。传入即可——`createPreview(el, { source: file, password })`。注意口令缺失或错误一律通过 `onError` 以 `PASSWORD_PROTECTED` 码返回，`createPreview` 自身从不抛异常，所以错误要在回调里处理，而不是 `try/catch`。示例见[基础用法](/zh/packages/excel-preview/examples/01-basic)。

### `.xls` / `.ods` 文件被直接拒绝

v1 有意如此，且每种情况有各自的错误码：旧版 `.xls`（BIFF8）为 `LEGACY_FORMAT`（请另存为 `.xlsx`）；`.ods` 等同样基于 ZIP 的办公套件为 `CORRUPT`（是没有 `xl/workbook.xml` 的 ZIP 包）；SpreadsheetML / HTML 表格导出为 `UNSUPPORTED`。UTF-16 的 CSV 同样落在 `UNSUPPORTED`——先另存为 UTF-8。完整对照表见[范围与限制](/zh/packages/excel-preview/guide/04-limits)。

### 能读 Excel 导出的中文 CSV 吗？

能。UTF-8（带或不带 BOM）与 GB18030 都会自动识别，分隔符也自动嗅探（`,`、`;`、制表符、`|`）。例外是 UTF-16 的 CSV（`FF FE` / `FE FF` BOM）——它会被判定为二进制并以 `UNSUPPORTED` 拒绝，请先另存为 UTF-8。

### `createPreview` 之后立刻调 `setSheet()` 没有生效

这是文档写明的时序，不是 bug：`createPreview` 是异步启动的，解析完成前 `setSheet()` 是静默空操作、`getSheetNames()` 返回 `[]`。请在 `onParsed` 里、或由你自己的 UI 调用，也可以直接用表格底部的工作表标签栏切换。另外 `destroy()` **不会**取消已在进行中的解析，它只是屏蔽该次解析本会触发的回调。详见[快速开始](/zh/packages/excel-preview/guide/01-quick-start)。

### 公式单元格渲染为空

公式单元格渲染的是文件中**缓存**的值——与 SheetJS、exceljs 同一惯例，预览从不计算公式。因此由非 Excel 写入器生成、且没有缓存值的文件会显示为空单元格；用 Excel 打开另存一次即可补上缓存。带缓存结果的公式会正常渲染（数值缓存为 `type: "number"`，字符串缓存为 `type: "string"`）。

### 预览大文件会卡死页面吗？

不会。解析在 Web Worker 内完成（实测 10 万行 × 10 列约 1.5s），渲染只挂载视口内的单元格，因此页面响应性与文件大小无关。但仍有两个真实上限：整个文件会在 worker 内一次性读入内存（数百 MB 的文件会先撞上内存上限）；表格是一个很高的容器元素，Firefox / Safari 在约 89 万行后无法继续滚动（Chrome 约 167 万行）。细节见[范围与限制](/zh/packages/excel-preview/guide/04-limits)。

## 进度遮罩（导出包的 `overlay` 选项）

### 很快结束的任务没有出现遮罩

这是 `delayMs` 的作用，默认 200ms：比它更早结束的任务完全不会显示遮罩，而不是只闪一帧。想立刻挂载就设 `delayMs: 0`。

### 长时间同步阻塞的任务里遮罩一直不出现

默认 `delayMs: 200` 时，长阻塞会"饿死"这次显示——延迟到期时主线程正忙，遮罩始终没能绘制。正确做法是提前挂载，并在阻塞调用前让出一帧：

```ts
import { nextPaint, showExportOverlay } from "@marcusok/excel-exporter/overlay";

const overlay = showExportOverlay({ delayMs: 0 });
try {
  await nextPaint(); // 先让浏览器把遮罩画出来
  heavySyncWork();
} finally {
  overlay.close();
}
```

`nextPaint()` 与 `showExportOverlay()` 都由 `@marcusok/excel-exporter` 的 `/overlay` 子路径导出。详见[阻塞主线程](/zh/packages/excel-exporter/guide/11-overlay#主线程阻塞)。

### 两个任务同时运行，显示的是谁的文案？

并发任务共用一个带引用计数的 DOM 节点，后写者胜：后调用的任务在 `showProgressOverlay` 时即接管显示，最后一个任务关闭后节点才移除。若某个提前退出的调用方，其尚未触发的显示定时器在它 close 之后才到期，显示的文案会归属于当前仍活跃的最近一个调用方。

### Node / SSR 下能用吗？

能。没有 `document` 时调用返回一个空操作句柄，因此同一处调用在两种环境下都成立——不需要分支判断，也不需要 `typeof window`。

## 引擎层（内置于两个包）

### 有需要单独安装的引擎包吗？

没有。引擎集成（WASM 加载器 + 对 modern-xlsx 的稳定再导出面）是仓库私有的内部层，构建期整体打进各业务包的 `dist`。两个包都**再导出**了它的 `configureWasm` / `getWasmLoader`——即便要自托管 WASM 二进制，也不必往你自己的依赖里加任何东西。

### `modern-xlsx` 要求 `engines.node >= 24`，会影响我的应用吗？

不会。引擎运行时在构建期已被**打进**各业务包的 `dist/`，消费方不会执行任何上游 modern-xlsx 代码。两个业务包自身要求 Node `>= 22`，浏览器侧需要支持 WebAssembly。manifest 里那条精确钉版的 `modern-xlsx` 依赖，只是为了 TypeScript 消费方能解析再导出的类型——运行时不加载它的任何代码。

### 导出包与预览包同在一页

各包各持一份内置引擎，因此该页面天然就是两个引擎实例、两份加载器状态——从其中一个包调 `configureWasm` 不会配置另一个包的 loader。WASM 的网络传输通常会被去重（同一份二进制、内容 hash 资产、HTTP 缓存）；内存开销真实存在，但只有两包真的同页运行时才需要在意。完整说明见[包关系与选型](/zh/guide/03-package-relationships)。
