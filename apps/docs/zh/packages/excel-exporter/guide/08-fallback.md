# 兜底机制（无样式快速流）

当 WASM 路径不可用时，库会自动降级到纯 JS 快速流，多数异常下仍能拿到导出文件（例外见下方触发条件末条）。兜底不依赖 WASM、不依赖 Worker、不依赖网络——它就是 ≥ 50,000 行导出使用的同一个流式写入器，在主线程以无样式模式运行。

## 触发条件

- 环境不支持 `WebAssembly`（仅影响 main 与 Worker + Workbook 路径；≥ 50,000 行的 stream 路径不依赖 WASM，不受此条影响）；
- `modern-xlsx.wasm` 加载失败（共尝试 `maxRetries` 次后仍失败，默认 3 次含首次）；
- Worker 路径失败（如 worker 资产 404）**且**自动的主线程重试也失败——Workbook 路由下重试会先在主线程运行 modern-xlsx（样式保留），流式兜底是最后保底，而非 Worker 失败后的直接下一步。≥ 50,000 行的 stream 路由下，重试本身就是同一个快速流（相同代码、相同输入），失败即终局，不会再做第三次尝试；
- 主线程 Workbook 路由的构建阶段抛错（如 modern-xlsx 构建期内部错误）。Node 上的 stream 路由（无 `window`，显式 `mode: "stream"` 或 auto ≥ 50,000 行）构建抛错则直接终局——首次尝试已经在同一份输入上跑过快速流，没有可重试或可兜底的路径。注意：结构性输入错误（非法/重复表名、越界合并区间、非数值或负数 `width`、非整数 `freezeRows`、越界 `format.decimals` 等）不进入此链路——前置校验在任何路由运行前即拦截，直接返回 `success: false`，不会发起兜底尝试。

## 行为差异

| 维度                  | modern-xlsx 路径 | 流式兜底                                                                                |
| --------------------- | ---------------- | --------------------------------------------------------------------------------------- |
| `ExportResult.engine` | `"modern-xlsx"`  | `"modern-xlsx"`，且 `mode: "stream"`、`result.error` 非空（`success` 仍为 `true`）      |
| 单元格样式            | 完整             | 剥离（快速流输出不带样式部件的极简 OOXML）                                              |
| 多行表头 / 合并       | 支持             | 支持（合并不是样式）                                                                    |
| 列宽/冻结/筛选        | 支持             | 不支持                                                                                  |
| FormatSpec            | 支持             | 支持（enum/padding/number/date 语义保留，日期输出为可读字符串）                         |
| 数字格式              | `numFormat`      | `decimals` 烧入存储值                                                                   |
| 警告                  | —                | console 打印 `[excel-exporter] Falling back to the style-less fast stream. Reason: ...` |

## 如何感知兜底

```ts
const result = await exportExcel(options);
if (result.success && result.error) {
  // 已降级：样式被剥离，原因在 result.error.message
}
```

兜底不是常规路径，而是异常时的保险。出现兜底请优先排查 wasm 资产是否 404（打开 Network 面板），以及使用 `configureWasm` 覆盖时 URL 是否在首次导出前配置。还有一种更隐蔽的变体：wasm 请求显示 **200/304 但响应是 HTML**（`content-type: text/html`；底层编译错误含 `expected magic word 00 61 73 6d, found 3c 21 64 6f`）——这是 Vite 开发服务器的依赖预构建落入了 HTML fallback，修复方法是在 `optimizeDeps.exclude` 中排除本包，详见[安装与配置 → Vite 开发服务器](/zh/packages/excel-exporter/guide/02-installation)。另请注意：首次尝试失败后，后续导出只会报 `WASM load previously failed` 而非原始错误——真实原因只出现在**第一次**导出（或刷新页面后首次导出）的 console 警告里。
