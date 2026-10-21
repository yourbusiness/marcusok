# 性能参考

以下数字为本机实测（真实 Chrome，6 列混合类型数据），用来说明自动路由的取舍依据。Node **进程内**回归基线（SLA 200/500/1000 ms）在 `src/__tests__/performance.test.ts`；toBuffer 断崖的独立进程验证见包 README——同为 100k Workbook，Node 冷启动实测约 21s，下文的 17.5s 是 Chrome 口径：同一个断崖、两种环境。

## 基准数据

| 数据量     | Workbook（main）  | Fast stream | auto 实际选择        |
| ---------- | ----------------- | ----------- | -------------------- |
| 10,000 行  | ~120ms            | —           | main + Workbook      |
| 50,000 行  | —                 | ~400ms      | Worker + Fast stream |
| 100,000 行 | 17.5s（历史基线） | ~780ms      | Worker + Fast stream |

<ClientOnly>
  <BenchmarkChart dir="excel-exporter" />
</ClientOnly>

宽表单独实测（Node 24.21）：40k×50 = 200 万格在 Workbook 路径约 6.3s / RSS +1.2GB，Fast stream 约 1s / +58MB；20k×100 = 200 万格 Workbook 约 3.2s / +258MB；60k×20 = 120 万格 stream 约 1s。因此路由同时看估算单元格数，而不只看行数。

## 结论

1. `Workbook.toBuffer()` 在 ~5.5 万行开始出现超线性断崖（10 万行 17.5s），而 Fast stream 约 0.8s，因此 stream 门槛是 5 万行或 150 万格；
2. 浏览器中行 ≥ 20,000 **或** 格 ≥ 250,000 时交给 Worker；主线程只做一次结构化克隆，其成本随单元格数增长（10 万×6 约 94ms、40k×50 约 200ms）；
3. Stream 路径的代价是单元格/布局样式缺失（数字格式已通过最小 `styles.xml` 保留），所以小文件默认走带完整样式的 Workbook 路径；`result.degradedFeatures` 让这次特性丢失可以被程序化感知。

## 优化建议

- 默认使用 `mode: "auto"`，不要手动指定 `main` 导出大文件；
- 需要样式时让**单次调用同时低于两个门槛**（行 < 5 万且格 < 150 万）——auto 按一次调用全部 sheet 的总行数/总单元格数路由，拆成同一次调用的多个 Sheet 没有用；更大数据量接受样式降级（检查 `result.degradedFeatures`），或拆成多次 `exportExcel` 调用（每次一个文件）；
- `decimals`/`thousands` 现在是两条路径共同的纯显示 numFmt 设置，无需为了跨阈值一致而预先舍入；
- 服务端大批量导出放到异步任务中执行。

> 运行时说明：本地基线与 `.nvmrc` 使用 Node 24.21.0。内置引擎的 WASM `readBuffer` 在 Node 22.x（以及 Node 24.14）上对数字密集表慢约 10 倍；导出的 Workbook/流式基准受影响较小，但性能对比应以锁定运行时为准。CI 另保留 Node 22 功能任务并跳过性能测试。
