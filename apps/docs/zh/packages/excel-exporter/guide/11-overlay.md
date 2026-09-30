# 进度遮罩

2.8.0 起，每次导出**默认**显示可配置的全屏遮罩——无需额外 import、无需包装函数。`overlay: false` 完全关闭；传配置对象则定制。遮罩 UI 已内置于本包（3.0 之前曾以独立包分发，通用入口现移至 `/overlay` 子路径），本页讲导出包如何驱动它。

```ts
import { exportExcel } from "@marcusok/excel-exporter";

const result = await exportExcel({
  filename: "sales-2026",
  sheets: [{ name: "销售", columns, data }],
  // overlay: false,            // <- 完全关闭
  // overlay: { delayMs: 0 },   // <- 定制
});
```

遮罩在短暂延迟后出现、阻断页面交互，并在导出结束时移除——**成功与失败都会移除**。Node/SSR 下是空操作（没有 `document`）。

`exportTable` 与 `exportEcharts` 虽然委托 `exportExcel`，但它们各自的选项类型（`TableExportOptions` / `EChartsExportOptions`）**不含** `overlay` 字段，其转换函数也不会透传该字段——在这里传 `overlay` 会 TypeScript 报错，运行时则被静默忽略，默认遮罩照常弹出。若想在这两种数据形态下控制遮罩，先用 `tableToSheet` / `echartsToSheet` 转换后再自行调用 `exportExcel`（见[表格与 ECharts 入口](/zh/packages/excel-exporter/api/05-table-and-echarts)）。

## 取值

| 值                       | 行为                                                                                                                                                                                     |
| ------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 缺省 / `true`            | 默认导出文案的遮罩（2.8.0 起的默认行为）。                                                                                                                                               |
| `false`                  | 完全没有遮罩：不挂载、不多让帧，回调行为与该特性出现之前完全一致。                                                                                                                       |
| `ProgressOverlayOptions` | 定制：文案、`delayMs`、主题、`blockInteraction` 等（全部字段见下方选项块）。定制文案与导出默认**合并**——覆盖 `text.title` 时内置阶段文案仍在，`building` 这类阶段 key 不会显示成裸 key。 |

```ts
await exportExcel({
  filename: "sales-2026",
  sheets: [{ name: "Sales", columns, data }],
  // `overlay` 是 ExportOptions 的字段——所有 ProgressOverlayOptions 条目
  // 都可以在这里定制。下例各值即内置导出默认值，只传想改的项即可。
  overlay: {
    delayMs: 200, // 导出在此毫秒内结束则完全不显示遮罩
    minVisibleMs: 300, // 已经显示过就至少停留这么久（避免一闪而过）
    fadeOutMs: 150,
    zIndex: 2147483000,
    container: document.body,
    blockInteraction: true, // 置 false 则只做视觉覆盖
    theme: "auto", // "auto" | "light" | "dark"
    text: {
      title: "正在导出 Excel",
      initial: "准备中…",
      phases: {
        building: "正在构建工作簿…",
        downloading: "正在下载…",
        finishing: "即将完成…",
      },
      hint: "数据量较大时可能需要数十秒，请勿关闭页面",
    },
  },
});
```

`text.hint` 只在进度条处于不确定态时渲染——确定态显示百分比，不再显示提示。

## 不确定态与确定态

进度来自库已有的回调（`onProgress` / `onPhase` 是**链式追加**而非替换——挂在同一对回调上的指标面板照常工作），所以遮罩的表现不会好于背后的数据源：

| 路由                 | 触发条件                                                    | 中间进度        |
| -------------------- | ----------------------------------------------------------- | --------------- |
| main + Workbook      | 浏览器 < 20,000 行                                          | 无              |
| main + Fast stream   | Node（auto ≥ 50,000 行，或显式 `stream`——Node 没有 Worker） | 每 1,000 行一次 |
| Worker + Workbook    | auto，20,000–49,999 行                                      | 无              |
| Worker + Fast stream | auto ≥ 50,000 行，或浏览器下的显式 `stream`                 | 每 1,000 行一次 |

只有 Fast stream 路径会在 `0` 与 `1` 之间上报 `onProgress`。因此遮罩在收到第一个中间值之前渲染的是**旋转圆环**，收到后才切成确定态进度条。Workbook 路由全程停留在不确定态——这是数据源的粒度决定的，不是渲染问题。

收尾的 `onProgress(1)` 既不会关闭遮罩，也不会在毫无真实进度的路由上伪造一条走完的进度条；关闭由 promise settle 驱动。

## 主线程阻塞

`WorkbookBuilder.addSheet` 与 Fast stream 写入器都是**同步**的。它们运行期间浏览器无法重绘，这直接约束了遮罩的行为：

- 默认 `delayMs: 200` 时，如果延迟到期那会儿阻塞已经开始，显示会被饿死：这次导出**根本不会出现遮罩**。这就是主路由上一次"小数据量、WASM 已热"的导出完全不弹遮罩的原因——构建在延迟到期前就开始并且早于它结束。迟到的定时器在关闭时会被显式清掉，所以它也绝不会在导出结束**之后**才弹出来。
- `delayMs: 0` 时遮罩会在构建**之前**同步挂载，再由两帧让出（`nextPaint`）把绘制机会让给浏览器。这是阻塞路由上唯一能看见遮罩的配置——代价是快导出会闪一下。

两种配置下，圆环在阻塞期间都会继续旋转（它是 CSS transform 动画，由合成器线程驱动），而百分比与文案会冻结到线程空闲为止。

## 自己驱动遮罩

围绕底层入口的自定义流程，经 `/overlay` 子路径直接驱动遮罩——导出包的 overlay 选项就是这套协议加导出语义文案：

```ts
import { showExportOverlay } from "@marcusok/excel-exporter/overlay";

const overlay = showExportOverlay({
  text: { title: "正在导出 Excel", phases: { building: "正在构建工作簿…" } },
});
try {
  // exportAsStream 的第二个参数是回调函数本身，不是选项对象
  return await exportAsStream(sheets, (p) => overlay.setProgress(p));
} finally {
  overlay.close(); // 幂等
}
```

## 历史子路径

`@marcusok/excel-exporter/overlay`（2.8 之前的历史入口）继续存在，且在遮罩不再单独发包之后，它就是导出流程之外驱动遮罩的正式途径：`exportExcelWithOverlay(options, overlay?)` 是等价于 `exportExcel({ ...options, overlay })` 的薄封装，`showExportOverlay` 是通用遮罩入口，`nextPaint` 也随之一并导出（用于[主线程阻塞](#主线程阻塞)中的先挂载再让帧写法）。注意随当年迁移文案结构有变：旧的平铺字段（`text.building`、`text.downloading`……）改成了 `text.phases` 表，句柄方法是 `setProgress` / `setPhase(key)`（不再是 `handleProgress` / `handlePhase`）。

## 并发

遮罩共用一个 DOM 节点并做引用计数：并发导出渲染进同一个遮罩（最后更新者为准），最后一个关闭时才移除。并发调用 `exportTable` / `exportExcel` 是安全的，多次运行之间不会互相残留。渲染内容（标题、提示、文案、进度）始终属于最近一次 show / progress / phase 事件的调用方。

## Node 与 SSR

没有 `document` 时遮罩是空实现句柄，`exportExcel` 照常导出——同一处调用代码在两种环境下都可用，无需分支。

## 无障碍

进度条带 `role="progressbar"`（仅在确定态下附带 `aria-valuenow`），文案节点是 `aria-live="polite"` 的状态区域，容器设置 `aria-busy`。`prefers-reduced-motion: reduce` 下会关闭圆环旋转动画与全部过渡。

需要留意的是，阻断只做在指针层面，遮罩**有意不**对下层页面施加 `inert`——键盘用户仍能 Tab 到视觉上被盖住的元素。把每个兄弟节点都置为 `inert` 侵入性太强，有破坏宿主页面行为的风险，因此没有采用。

## 相关内容

- [自己驱动遮罩](#自己驱动遮罩)——`/overlay` 子路径入口，用于导出之外的流程。
- [进度与阶段回调](./10-advanced#进度与阶段回调)——底层的 `onProgress` / `onPhase` 契约。
- [Worker 与 stream 模式](./06-worker-stream)——给定行数会走哪条路由。
