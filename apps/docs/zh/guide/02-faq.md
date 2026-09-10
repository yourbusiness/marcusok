# 常见问题

### 浏览器报 WASM 404

默认（零配置）定位下，Vite / webpack 5 会把随包发布的 `modern-xlsx.wasm` 自动发射为 hash 资产，不应出现 404。若仍遇到，通常是覆盖了 URL（`configureWasm({ wasmUrl })` 指向了错误路径），或使用的打包器不支持 `new URL(资产, import.meta.url)` 资产模式；把 `configureWasm` 指向站点实际可访问的地址，或从本包 `dist/` 把文件拷贝到静态目录即可。

### Worker 模式回退到了主线程

worker 资产（`export.worker.js`）默认自动定位；回退发生在 Worker 路由失败时（例如 `workerUrl` 覆盖配置指向了 404 的地址）。此时导出会**在主线程重试**（modern-xlsx 保留样式；Fast stream 本身不需要 WASM）——只有主线程重试也失败时，才最后降级到无样式的流式兜底。查看 console 中 `[excel-exporter]` 前缀的警告可定位原因。

### 导出成功但 result.error 有值

说明导出降级到了无样式快速流（WASM 失败或环境不支持）：样式被剥离，表头与合并保留。原因见 `result.error.message` 与 console 中的 `[excel-exporter]` 警告——通常是 wasm 资产 404。详见 [兜底机制](/zh/packages/excel-exporter/guide/08-fallback)。

### 10 万行数据导出非常慢（>15s）

大概率走了 `main` + `Workbook.toBuffer()` 路径——该路径在 ~5.5 万行后出现性能断崖。把 `mode` 保持为 `auto`（10 万行约 0.8s），或显式指定 `mode: "stream"` / `mode: "worker"`。详见 [自动模式路由](/zh/packages/excel-exporter/guide/03-auto-mode)。

### Stream 模式下样式不生效

Stream 路径 v1 支持多行表头（`children`）与数据区合并（`merges`），但不支持单元格样式、表头样式与列宽/冻结/筛选等布局特性（会在 console 打印警告）。需要完整样式时，控制在 5 万行以内走 Workbook 路径。详见 [Worker 与流式](/zh/packages/excel-exporter/guide/06-worker-stream)。

### 导出是本地完成的吗？

是。所有处理都在浏览器/Node 进程内完成，不上传任何业务数据。

### 日期列显示为长文本，Excel 不识别为日期

不声明 `format` 时，`Date` 值会按普通文本写入单元格，Excel 不会识别为日期：所有路径（main / worker / stream）统一写入 ISO 字符串（如 `2026-07-01T00:00:00.000Z`）。日期列需要声明 `format: { type: "date" }`（或 `datetime`）：Workbook 路径会写入 Excel 日期序列并自动注入对应 `numFormat`，单元格才会被 Excel 识别为真正的日期。
