# 进度遮罩

导出期间显示的**可选**全屏遮罩 + 进度条。它挂在独立子路径下，因此主入口（`@marcusok/excel-exporter`）不含任何 DOM 代码，不用这个能力的调用方不会多打包一个字节。

```ts
import { exportExcelWithOverlay } from "@marcusok/excel-exporter/overlay";

const result = await exportExcelWithOverlay({
  filename: "sales-2026",
  sheets: [{ name: "销售", columns, data }],
});
```

遮罩在短暂延迟后出现、阻断页面交互，并在导出结束时移除——**成功与失败都会移除**。

## 配置项

```ts
await exportExcelWithOverlay(options, {
  delayMs: 200, // 导出在此毫秒内结束则完全不显示遮罩
  minVisibleMs: 300, // 已经显示过就至少停留这么久（避免一闪而过）
  fadeOutMs: 150,
  zIndex: 2147483000,
  container: document.body,
  blockInteraction: true, // 置 false 则只做视觉覆盖
  theme: "auto", // "auto" | "light" | "dark"
  text: {
    title: "正在导出 Excel",
    preparing: "准备中…",
    building: "正在构建工作簿…",
    downloading: "正在下载…",
    finishing: "即将完成…",
    hint: "数据量较大时可能需要数十秒，请勿关闭页面",
  },
});
```

| 配置项             | 默认值          | 说明                                                        |
| ------------------ | --------------- | ----------------------------------------------------------- |
| `delayMs`          | `200`           | 挂载遮罩前的延迟。比它更快的导出根本不会显示遮罩。          |
| `minVisibleMs`     | `300`           | 一旦显示就至少停留这么久——移除走延时，而不是一闪而过。      |
| `fadeOutMs`        | `150`           | 节点移除前的淡出时长。                                      |
| `zIndex`           | `2147483000`    | 遮罩层级。                                                  |
| `container`        | `document.body` | 挂载容器。                                                  |
| `blockInteraction` | `true`          | 在遮罩上拦截指针与滚动事件；置 `false` 时下层页面仍可操作。 |
| `theme`            | `"auto"`        | `"auto"` 在挂载时按 `prefers-color-scheme` 解析。           |
| `text`             | 中文            | 文案覆盖；不确定态下会额外显示 `hint`。                     |

`text.hint` 只在进度条处于不确定态时渲染——确定态显示百分比，不再显示提示。

## 不确定态与确定态

进度来自库已有的回调，所以进度条的表现不会好于背后的数据源：

| 路由                 | 触发条件                                                    | 中间进度        |
| -------------------- | ----------------------------------------------------------- | --------------- |
| main + Workbook      | 浏览器 < 20,000 行                                          | 无              |
| main + Fast stream   | Node（auto ≥ 50,000 行，或显式 `stream`——Node 没有 Worker） | 每 1,000 行一次 |
| Worker + Workbook    | auto，20,000–49,999 行                                      | 无              |
| Worker + Fast stream | auto ≥ 50,000 行，或浏览器下的显式 `stream`                 | 每 1,000 行一次 |

只有 Fast stream 路径会在 `0` 与 `1` 之间上报 `onProgress`。因此遮罩在收到第一个中间值之前渲染的是**流动扫光**，收到后才切成确定态进度条。Workbook 路由全程停留在不确定态——这是数据源的粒度决定的，不是渲染问题。

收尾的 `onProgress(1)` 既不会关闭遮罩，也不会在毫无真实进度的路由上伪造一条走完的进度条；关闭由 promise settle 驱动。

## 主线程阻塞

`WorkbookBuilder.addSheet` 与 Fast stream 写入器都是**同步**的。它们运行期间浏览器无法重绘，这直接约束了遮罩的行为：

- 默认 `delayMs: 200` 时，如果延迟到期那会儿阻塞已经开始，显示会被饿死：这次导出**根本不会出现遮罩**。迟到的定时器在关闭时会被显式清掉，所以它也绝不会在导出结束**之后**才弹出来。
- `delayMs: 0` 时遮罩会在 `exportExcel` 调用**之前**同步挂载，再由两帧让出把绘制机会让给浏览器。这是阻塞路由上唯一能看见遮罩的配置——代价是快导出会闪一下。

两种配置下，扫光动画在阻塞期间都会继续流动（它是 CSS transform 动画，由合成器线程驱动），而百分比与文案会冻结到线程空闲为止。

## 句柄式用法

用于 `exportTable`、`exportEcharts` 或自定义流程时，自己驱动遮罩：

```ts
import { showExportOverlay } from "@marcusok/excel-exporter/overlay";

const overlay = showExportOverlay({ delayMs: 200 });
try {
  return await exportTable({
    columns,
    data,
    filename: "报表",
    onProgress: (p) => overlay.handleProgress(p),
    onPhase: (phase, ms) => overlay.handlePhase(phase),
  });
} finally {
  overlay.close(); // 幂等
}
```

`exportExcelWithOverlay` 就是这个模式的封装：它对调用方已有的 `onProgress` / `onPhase` 是**链式追加**而非替换，因此挂在同一对回调上的指标面板照常工作。

## 并发

遮罩共用一个 DOM 节点并做引用计数：并发导出渲染进同一个遮罩（最后更新者为准），最后一个关闭时才移除。并发调用 `exportTable` / `exportExcel` 是安全的，多次运行之间不会互相残留。

## Node 与 SSR

没有 `document` 时，`showExportOverlay` 返回空实现句柄，`exportExcelWithOverlay` 照常导出——同一处调用代码在两种环境下都可用，无需分支。

## 无障碍

进度条带 `role="progressbar"`（仅在确定态下附带 `aria-valuenow`），文案节点是 `aria-live="polite"` 的状态区域，容器设置 `aria-busy`。`prefers-reduced-motion: reduce` 下会关闭扫光动画与全部过渡。

需要留意的是，阻断只做在指针层面，遮罩**有意不**对下层页面施加 `inert`——键盘用户仍能 Tab 到视觉上被盖住的元素。把每个兄弟节点都置为 `inert` 侵入性太强，有破坏宿主页面行为的风险，因此没有采用。

## 相关内容

- [进度与阶段回调](./10-advanced#progress-and-phase-callbacks)——底层的 `onProgress` / `onPhase` 契约。
- [Worker 与 stream 模式](./06-worker-stream)——给定行数会走哪条路由。
