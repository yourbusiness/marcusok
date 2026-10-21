# 自动模式路由

`exportExcel` 的 `mode` 参数默认是 `"auto"`：库根据**总行数与估算单元格数**以及运行环境自动选择路径，业务代码无需感知数据量变化。单元格数 = 数据行数 × 有效叶子列数（`indexColumn` 展开后），按一次调用中的全部 sheet 累加。

## 路由规则

### 浏览器

| 数据量                                                                   | 路径                   | 说明                                                                  |
| ------------------------------------------------------------------------ | ---------------------- | --------------------------------------------------------------------- |
| `< 20,000` 行**且** `< 250,000` 格                                       | `main`                 | 主线程同步构建，10k×6 列浏览器实测约 120ms                            |
| 行 `>= 20,000` **或** 格 `>= 250,000`，且行 `< 50,000`、格 `< 1,500,000` | `worker` + Workbook    | 主线程只做一次结构化克隆（成本随单元格数增长），WASM 在 Worker 内执行 |
| 行 `>= 50,000` **或** 格 `>= 1,500,000`                                  | `worker` + Fast stream | 自研 fflate 写入，避免 `toBuffer` 性能断崖                            |

### Node / SSR（无 Web Worker）

| 数据量                                  | 路径     | 说明                               |
| --------------------------------------- | -------- | ---------------------------------- |
| 行 `< 50,000` **且** 格 `< 1,500,000`   | `main`   | 主线程 Workbook 构建，保留完整样式 |
| 行 `>= 50,000` **或** 格 `>= 1,500,000` | `stream` | Fast stream，10 万行约 0.8s        |

## 手动指定

```ts
await exportExcel({ ..., mode: "stream" });  // 强制流式
await exportExcel({ ..., mode: "worker" });  // 强制 Worker（浏览器）
await exportExcel({ ..., mode: "main" });    // 强制主线程
```

`mode: "worker"` 在 Node/SSR 环境不会报错：库会回退到主线程路径（达到 stream 门槛后用 stream），保证样式语义不丢失，而不是降级成无样式的流式路径。浏览器内显式 `worker` 模式同样用行/单元格门槛决定 Worker 内走 Workbook 还是 Fast stream。

## 阈值为什么是 20,000 / 50,000 / 单元格预算

- **20,000 行或 250,000 格**：主线程同步工作在此量级以下可接受；超过后浏览器交给 Worker，避免阻塞 UI。
- **50,000 行或 1,500,000 格**：实测 `Workbook.toBuffer()` 在 ~5.5 万行开始出现超线性性能断崖（10 万行 Chrome 实测 ~17.5s，Node 冷启动约 21s，见[性能页](/zh/packages/excel-exporter/guide/07-performance)），而 Fast stream 约 0.8s。单元格门槛补上了另一个盲区：宽表可能远低于 5 万行，但 Workbook 路径依然很重。Node 实测：40k×50（200 万格）Workbook ~6.3s / RSS +1.2GB，stream ~1s / +58MB；20k×100（200 万格）Workbook ~3.2s / +258MB；60k×20（120 万格）stream ~1s。

> 代价说明：Stream 路径支持多行表头、单元格合并与数字格式（date/number FormatSpec 经最小 `styles.xml`），但不支持单元格样式、表头样式与列宽/冻结/筛选等布局特性。auto 模式按**一次调用中全部 sheet 的总行数与总单元格数**路由，把数据拆成同一次调用的多个 sheet 并不能保住样式。需要完整样式的导出请让单次调用同时低于两个门槛，或拆成多次 `exportExcel` 调用（每次一个文件）。
>
> Stream 路径丢弃特性时，`ExportResult.degradedFeatures` 会列出具体项（`width`、`headerStyle`、`style`、`dataStyle`、`freezeRows`、`autoFilter`）。正常的大文件 stream 路由即使 `success: true`、`error` 为空也会设置该字段；失败驱动的降级还会同时设置 `error` 说明原因。
