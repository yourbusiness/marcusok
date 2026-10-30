# 限制与安全说明

## 渲染

- **懒渲染管线**——页面进入 1 屏预载带即渲染，离开 3 屏保留带即释放 canvas。快速滚动长文档时只渲染最终停留的页；已入队但不再需要的页会在开渲前被丢弃。
- **单页串行**——每实例同时最多渲染一页（滚动洪峰不会并发引爆 canvas 工作），滚动交错时按观察器上报顺序处理。
- **占位精度**——每页占位按该页自身 viewport 计算，混合尺寸文档（如 US Letter + A4 横向）不会出现滚动条跳动。全部页代理在挂载后立即预取，滚动总高一次成型。
- **清晰度**——canvas 按 `devicePixelRatio` 分辨率绘制、CSS 尺寸占位，与官方 viewer 同款。
- **缩放只解析一次**——`fit-*` 在 create 时量测容器。`fit-page` 需要可量测且高度受约束的容器（定高、`max-height` 或 flex 解析出的高度）；高度随内容增长的容器里量到的不是视口高度，请给定高或改用 `fit-width`。容器尺寸变化不会自动重排。
- **滚动条时机**——`fit-width` 的量测发生在页占位存在**之前**，即滚动容器出现竖向滚动条之前。在经典（非 overlay）滚动条平台上，内容盒随后收窄一个滚动条宽（约 15–17px），居中的页面可能横向溢出这一宽度。重新 create 无济于事（空态下重新量测复现同样问题）。在意的话给容器加 `scrollbar-gutter: stable` 或在布局中预留 gutter。
- **`devicePixelRatio` 按渲染时刻取值**——跨屏拖动窗口到不同 DPR 屏幕（或缩放）不会重渲已渲染的页；已渲染页面保持旧位图清晰度，直到被释放后重渲。

## 格式

- `%PDF-` 魔数必须出现在**头部 1024 字节内**（Adobe 自家阅读器口径）。Adobe 规范允许文件头前有垃圾字节，官方 viewer 搜索窗口更宽——垃圾前缀超 1 KB 的极端文件在官方 viewer 可开、在本包归 `UNSUPPORTED`。这是有意的输入收紧。
- **加密 PDF** 被识别并上报 `PASSWORD_PROTECTED`（message 区分上游"需要密码"与"密码被拒"两种 code）。首版不解密。
- **XFA 表单**不渲染（`enableXfa` 固定 `false`）；纯 XFA 文档显示空白页。
- **不支持 Node 服务端调用**：本包打包的 pdf.js 现代构建在 Node 下 `getDocument` 崩溃。import 安全；调用 `createPdfPreview` 需要浏览器。服务端文本提取不在范围内。

## Worker 回退（静默降级）

解析通常在 pdf.js Web Worker 内进行，但当 worker 无法构造时（CSP `worker-src` 拦截、MIME 类型错误、宿主对 module worker 的限制），pdf.js 自身会**静默回退到主线程 fake worker**。该模式下解析照常工作（最多在控制台留一条 `Setting up fake worker.` 警告），但全部发生在主线程：长文档会造成 UI 卡顿。两点须知：

- pdf.js 的"worker 已禁用"标志对**页面会话是粘性的**——任何一次构造失败后，页面上的后续实例（来自本包或任何其他 pdf.js 副本）都走 fake worker，直到刷新页面。
- 只有连主线程 import worker 模块都失败时，本包才会以提及 `workerSrc` 的 `UNKNOWN` 错误上抛。需要 worker 隔离时，请在 DevTools 确认存在 `pdf.worker` 线程，而不是默认渲染成功即 worker 生效。

## 安全加固（本包做了什么）

- 引擎运行**之前**先嗅探：非 PDF 输入得到可读的 `UNSUPPORTED`，而不是 worker 内部的模糊失败。
- **内嵌 PDF JavaScript 永不执行**：pdf.js 6.x 把脚本执行移入 sandbox 模块（QuickJS），本包从不加载该模块。（旧的 `isEvalSupported` 开关已在上游移除。）
- **注释层**（超链接、启动动作、表单控件）首版不渲染，收窄了主要的交互式攻击面。注释**外观**（已扁平化的盖章、水印）仍会随画布渲染——这是上游渲染的默认行为；首版跳过的是交互层。
- 渲染产物是 canvas 像素 + DOM 文本节点——没有 iframe、没有 `srcdoc`、没有 HTML 注入路径。

## wasm 门控特性（不随包分发）

pdf.js 的 wasm 资产（`jbig2.wasm`、`openjpeg.wasm`、`qcms_bg.wasm`）与 ICC 配置有意**不**再分发。未传 `wasmUrl` 时：

| 内容                 | 无 wasm 时的行为                    |
| -------------------- | ----------------------------------- |
| JPEG2000（JPX）图像  | JS 兜底解码器（较慢），不可用则跳过 |
| JBIG2 图像（扫描件） | JS 兜底解码器（较慢），不可用则跳过 |
| ICC 色彩管理图像     | 不做色彩管理直接渲染（上游警告）    |

要恢复完整保真，从 `pdfjs-dist` npm 包自托管 `wasm/`（及可选 `iccs/`）目录并传 `wasmUrl` / `iccUrl`。wasm 缺失**静默**降级——这也是要用真实文档做冒烟验证的又一原因。

## CSP 要求

- `font-src data:` —— pdf.js 以 `data:font/...` URL 注入内嵌字体。
- `style-src`（样式表层）—— pdf.js 的 FontLoader 运行时向 `adoptedStyleSheets` 或 `<head><style>` 注入 `@font-face` 规则。这是上游行为，**不在**本包 `styleNonce` 的覆盖范围。
- `style-src-attr 'unsafe-inline'` —— 文本层 span 与页占位使用内联 `style` 属性（`styleNonce` 只覆盖本包自己注入的样式表）。
- `worker-src`/`script-src` —— pdf.js worker 是同源（或你指定的 `workerSrc` 位置）加载的 module worker；跨源 `workerSrc` 会让 pdf.js 用 blob worker 包一层 import。

## 内存

- 被释放的页会取消在途文本层流、将 canvas 位图清零（`width = height = 0`）、移除文本层 DOM，并调用 pdf.js 的 `page.cleanup()` 释放主线程 operator list / objs；重新进入视口时按需重渲（不重复解析 PDF）。
- `destroy()` 还会取消在途渲染任务与文本流、清零已渲染 canvas，并在摘除前清空 wrapper。
- 文档代理与 worker 存活到 `destroy()` 为止；从 DOM 移除预览但不 `destroy()` 会泄漏 worker。文档级加载失败（`onError` 路径）会自行立即终止 worker——只有成功挂载的预览才会持有 worker。
- 全页代理预取（占位尺寸需要）会让每页代理在文档生命周期内保留——这是 pdf.js `getPage` 缓存的正常行为；canvas 位图与 operator list 不会累积，因为释放时会清理。

## SSR

import 安全（模块顶层不求值 DOM 全局），调用不安全：`createPdfPreview` 立即触碰 `document`/`IntersectionObserver`。与其他 DOM 库一样放进 `onMounted`/`useEffect` 守卫。在 Node 下 import 本包会让 pdf.js 现代构建打印一条 `Please use the legacy build in Node.js environments` 警告——属预期行为，并不代表服务端可以渲染。

## 多实例

实例在各自容器内完全隔离。注入的样式表（页面布局 + 文本层）是文档级幂等单例、全部实例共享，且有意在 `destroy()` 后保留。

- pdf.js 的 `GlobalWorkerOptions.workerSrc` 是模块级全局：多实例传入不同 `workerSrc` 时，后创建的实例生效。同页多实例请使用同一 worker URL（或默认推导）。
- 样式单例只注入一次，首个实例的 `styleNonce` 生效；多实例请使用相同 nonce，或都不传。
