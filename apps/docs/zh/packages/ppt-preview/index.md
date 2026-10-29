# @marcusok/ppt-preview

浏览器端 `.pptx` 只读预览 —— 连续滚动 + 逐页懒渲染 Canvas、可选中复制的文本层、0-based 页导航与加密文件密码解密。渲染引擎为 Rust/WASM 解析器（[@silurus/ooxml](https://github.com/yukiyokotani/office-open-xml-viewer) 0.89.6，构建期打进本包 `dist` —— 消费方零传递依赖），外层包了格式嗅探、source 防御性复制与视口驱动的显存管理。

## 能力

| 能力        | 说明                                                                                                                                                                      |
| ----------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 格式        | 标准 OOXML `.pptx`。带打开密码的文件（OLE2/MS-OFFCRYPTO **Agile** 加密）提供 `password` 即可解密。老式二进制 `.ppt` 被识别并报 `LEGACY_FORMAT`（不会静默错误渲染）        |
| 懒渲染      | 连续滚动；每页按 deck 宽高比定尺寸占位（无滚动条跳动）。页面进入视口 ±1 屏才渲染，远离视口 ±3 屏外自动释放 canvas 显存                                                    |
| 文本层      | 文本 runs 在正式渲染的**同一趟**收集（`onTextRun`，无双倍渲染）；透明文本层与 canvas 像素对齐 —— 可像真实文本一样选中复制。`enableTextSelection: false` 一行关闭          |
| 导航        | `goToSlide(index)` / `nextSlide()` / `prevSlide()`，0-based，基于当前最可见页跟踪（手动滚动始终被尊重）。隐藏页全部显示并计数，保持原始 deck 页序                         |
| Worker 解析 | 解析在 module Web Worker（WASM）内完成；幻灯片排版与 Canvas 绘制在主线程（上游 `mode: "main"` 的固定形态）                                                                |
| 解密        | MS-OFFCRYPTO Agile（SHA-512 密钥派生 + AES-CBC），经主线程 WebCrypto 执行 —— 需要安全上下文（HTTPS/localhost）。通常在解析前增加约 1s；ZIP 条目加密不在覆盖范围           |
| 加固默认值  | 引擎运行前先嗅探（空输入 / 非 ZIP 非 OLE2 提前报错；OLE2 放行给引擎区分 legacy/加密/未知 CFB）。超链接 inert（仅选择/复制）；媒体渲染为静态 poster（`skipMediaControls`） |
| 取消        | `destroy()` 断开观察器、作废在途渲染、canvas 置 0、移除注入 DOM、销毁引擎（终止解析 Worker）；所有回调静默。取消不是错误（不触发 `onError`）                              |
| 错误        | 携带错误码：`UNSUPPORTED`、`LEGACY_FORMAT`、`PASSWORD_PROTECTED`、`CORRUPT`、`UNKNOWN`                                                                                    |

## 安装

```bash
pnpm add @marcusok/ppt-preview
```

引擎（约 2.5 MB JS）与解析器 wasm（约 2 MB，gzip 后约 800 KB）随包打进 `dist`。Vite/webpack 下 wasm **零配置** —— bundler 检测到入口旁的资产即自动 emit。消费方构建可能出现两条关于 worker 资产（`pptx-source-worker.mjs`、`assets/render-worker-*.js`）的警告，属预期且可安全忽略：它们属于引擎 `mode: "worker"` 渲染路径与 legacy 路径，本包永不加载。见[边界说明](/zh/packages/ppt-preview/guide/02-limits)。

## 最小示例

```ts
import { createPptPreview } from "@marcusok/ppt-preview";

const preview = createPptPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  slideWidth: 960, // CSS 像素（默认）；高度按 deck 比例
  password: "…", // 可选；仅 MS-OFFCRYPTO Agile 文件
  onRendered: ({ slideCount, firstSlideMs, durationMs }) =>
    console.log(slideCount, firstSlideMs, durationMs),
  onError: (e) => console.error(e.code, e.message),
});

preview.nextSlide(); // 0-based 导航
preview.destroy(); // 取消在途工作 + 终止引擎
```

本包有意不提供底层数据模型 API：引擎的幻灯片模型绑定在活跃解析会话上，无法诚实地剥离成纯数据的 `parse()` 结果。这一点与提供 `parseWorkbookBytes` 的 [@marcusok/excel-preview](/zh/packages/excel-preview/) 不同，该模式见其文档。

<ClientOnly>
<PptPreviewDemo />
</ClientOnly>

## 首版不做（Not in scope）

- **老式 `.ppt`** —— 识别并报 `LEGACY_FORMAT`；请先转换为 `.pptx`。
- **动画、切换、音视频播放** —— 纯静态预览；媒体保留 poster/图形表现（不出现貌似可操作的播放控件）。
- **超链接** —— 渲染为文本但不可点击（inert）；外链 scheme 白名单与内部页跳转作为后续项。
- **容器重排** —— `slideWidth` 在创建时一次定型；容器尺寸变化不重排（destroy 后重建即可）。
- **可选引擎渲染器** —— MathJax（OMML 公式）、TIFF 图像、ChartEx、3-D 图表与 Region Map 不注入；受影响内容按引擎降级（占位符 / 2-D 家族）。见[边界说明](./guide/02-limits)。
