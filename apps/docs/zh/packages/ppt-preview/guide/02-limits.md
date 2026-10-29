# 边界与限制

本包有意不做的事与全部环境约束。提 issue 前请先读这一页 —— 多数"bug"都是这里写明的取舍。

## 格式支持

- **不支持老式二进制 `.ppt`。** 引擎确认 legacy 容器后本包报 `LEGACY_FORMAT`，请先转换为 `.pptx`。
- **密码解密只覆盖 OLE2/MS-OFFCRYPTO Agile**（PowerPoint 的「打开密码」，现代默认格式）。ZIP 条目加密不覆盖、不承诺。
- 解密的必知细节：
  - SHA-512 密钥派生与 AES-CBC 在**主线程**经 WebCrypto 执行 —— 需要安全上下文（HTTPS / localhost），通常在解析前增加**约 1s**（由 spinCount 决定）。
  - 上游解密**不校验 HMAC 数据完整性标签**；密文被篡改不会通过该机制被发现。

## 加载语义 —— 懒渲染 ≠ 懒解析

引擎的非渐进加载在 `load()` resolve 前**会预检（preflight）全部幻灯片**。本包的懒渲染只延迟 Canvas 绘制与文本层；首屏耗时仍随总页数增长。`onRendered.durationMs` 因此包含全 deck preflight（回调的文档注释中也写明了这一点）。

**加载期间的瞬态内存**：输入字节会被防御性复制两份 —— 一份用于嗅探并留存，另一份独占 buffer 交给解析器（解析器会把它 transfer 进自己的 worker）。因此加载期间源数据相关的峰值内存约为文件大小的 2 倍（瞬态）；调用方的原始字节全程不受影响，加载结束后副本即可被回收。

## 媒体、超链接与交互

- **纯静态预览**：动画、切换、音视频播放不渲染。媒体元素保留静态 poster/图形表现；显式跳过媒体播放控件（`skipMediaControls`），poster 上不会出现貌似可操作的控件。
- **超链接 inert**（首版）：链接文本正常渲染（可选择），但点击无动作。安全开放链接需要外链 scheme 白名单与内部页跳转设计（后续项）。
- **隐藏页全部显示**，计入 `slideCount` 并参与导航 —— 索引保持原始 deck 页序。无 show/skip/dim 配置。

## 内容降级（未注入可选引擎渲染器）

本包有意不注入引擎的可选渲染器模块。受影响内容按以下方式降级：

| 内容          | 无渲染器时的降级形态                                                                                                             |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------- |
| OMML 公式     | 回退表现（不加载 MathJax —— 无网络字体请求）                                                                                     |
| TIFF 图像     | 识别到的 TIFF 部件报渲染错误（不走浏览器解码器）                                                                                 |
| ChartEx 图表  | 标准的"不支持图表"占位符                                                                                                         |
| 3-D 图表组    | 回退到对应的 2-D 图表家族                                                                                                        |
| Region Map    | 标准的"不支持图表"占位符                                                                                                         |
| SmartArt      | 有 PowerPoint 已保存的 `dsp:drawing` 时按其渲染；缺失时引擎退为纯文本列表 —— 没有原生 diagram 布局引擎，不还原原始 SmartArt 布局 |
| 经典 2-D 图表 | 开箱支持（包含在格式入口内）                                                                                                     |

## 字体与 CSP

- **内嵌字体**正常使用，destroy 时由引擎释放。Google Fonts **默认关闭**（`useGoogleFonts: false` —— 不发起任何隐式外部请求）；开启后可能从 `fonts.googleapis.com` 拉取替代字形（CSP 需放行该域名；隐私上有外部请求）。
- **CJK 回退**：`renderOptions.cjkFallback`（默认 `"auto"`）为缺失的 CJK 字体选择替代字形（`sc`/`tc`/`hk`/`jp`/`kr`）。
- CSP 要求：
  - `worker-src blob:` 与 `worker-src data:` —— 引擎经 Blob URL 内联解析 worker（data: 兜底）。opaque-origin iframe（无可用 origin 的沙箱）可能无法创建 worker。
  - WASM 执行需要 `script-src 'wasm-unsafe-eval'`（或等价配置）。
  - `style-src-attr 'unsafe-inline'` —— 引擎文本层节点携带内联 `style` 属性；`styleNonce` 只覆盖**本包**注入的样式表，不覆盖引擎的内联样式。
  - CSP 无法放行内联 style 属性时，设置 `enableTextSelection: false`（纯 canvas）。

## Worker 与多实例内存

- 每个实例独占一个解析 Worker 与一份 WASM 实例化。同页多实例按倍数叠加该成本 —— 例如主预览 + 缩略图。每实例的主线程渲染是串行的，但各解析 Worker 并发运行。
- 消费方 bundler 可能打印的两条关于 `pptx-source-worker.mjs` / `assets/render-worker-*.js` 的构建警告，指向引擎 `mode: "worker"` 渲染路径与 legacy 路径的资产 —— **本包永不加载它们**（固定 `mode: "main"`）。警告可安全忽略；消除它们反而需要随包多发约 2 MB 死资产。

## 资源预算

超出引擎压缩包预算（条目大小、解压总量、条目数量、解码图像预算）的**合法文件**报 `UNSUPPORTED` —— **不是**文件损坏。经 `renderOptions.resourceLimits` 调整（`maxArchiveEntryBytes`、`maxTotalInflatedBytes`、`maxArchiveEntries`），见 [API](/zh/packages/ppt-preview/api/01-create-ppt-preview) 页。

## 布局与尺寸

- `slideWidth` 创建时定型；**无 resize 重排** —— destroy 后重建。
- 空 deck（`slideCount = 0`）加载不报错，也不触发 `onRendered`。
- SSR：模块可在服务端 import；`createPptPreview` 只能在存在 `Worker` 与 WebCrypto 的环境执行。

## 错误归一边界

- 引擎的 `not-ooxml`（非 pptx 的 zip、不可读容器、部分损坏的加密容器）归一为 `UNSUPPORTED` —— message 原样保留，但**不**进一步细分损坏原因（上游错误没有可依据的 `stage` 字段）。
- 未知解析陷阱、Worker/WASM/CSP 与资产 URL 失败归 `UNKNOWN` —— 绝不误报为文件损坏。
