# 在线演示

本页所有演示都在浏览器内直接运行发布态的库包——解析、构建与进度全在本地完成，不上传任何数据。数据由文档站的确定性 mock 生成器产生（mulberry32 种子 PRNG），同一配置多次生成结果一致，方便复现。

## Excel 导出 —— `@marcusok/excel-exporter`

选择数据集、数据量与导出模式，点击导出即可得到真实 `.xlsx` 文件，同时展示进度、各阶段耗时与最终引擎信息。

<ClientOnly>
  <PackageDemo dir="excel-exporter" />
</ClientOnly>

- **auto（推荐）**：按数据量自动选择 main / worker / stream 最优路径；
- **main**：主线程同步构建，10 万行时可以看到明显的性能断崖；
- **worker**：Web Worker 多线程，主线程只做一次结构化克隆；
- **stream**：Fast stream，10 万行约 0.8s，但不支持样式与布局特性。

完整用法见 [excel-exporter 文档](/zh/packages/excel-exporter/)。

## xlsx 预览 —— `@marcusok/excel-preview`

把 `.xlsx` / `.xlsm` / `.csv` 文件拖入面板：解析在 Web Worker 内完成，表格走 DOM 虚拟滚动渲染，大文件下面也不会卡顿。文件解析完成后会给出各阶段耗时（解析 / 渲染 / 合计）；加密文件会明确报 `PASSWORD_PROTECTED`，不会静默失败。

<ClientOnly>
  <PackageDemo dir="excel-preview" />
</ClientOnly>

完整用法见 [excel-preview 文档](/zh/packages/excel-preview/)。

## Word 预览 —— `@marcusok/word-preview`

把 `.docx` / `.docm` 文件拖入面板：文档按 Word 存盘的样子视觉保真渲染（分页、页眉页脚、脚注），每次渲染的完整耗时经 `onRendered` 上报。渲染在主线程完成（引擎约束——没有 Worker 模式）；不可信输入的加固默认开启（altChunk 关闭、超链接 scheme 白名单）。面板里内置了 1000 段落合成文档与双实例隔离演示。

<ClientOnly>
  <PackageDemo dir="word-preview" />
</ClientOnly>

完整用法见 [word-preview 文档](/zh/packages/word-preview/)。

## PDF 预览 —— `@marcusok/pdf-preview`

把 `.pdf` 文件拖入面板：解析在 pdf.js Web Worker 内完成，页面滚动进入视口才懒渲染（远离视口的页自动释放 canvas 显存），首页出图与端到端耗时经 `onRendered` 上报。透明文本层让预览内容可选中、可复制；加密 PDF 上报 `PASSWORD_PROTECTED`。本站的 CMap / 标准字体目录由 `public/` 提供（bundler 只 emit worker）。

<ClientOnly>
  <PackageDemo dir="pdf-preview" />
</ClientOnly>

完整用法见 [pdf-preview 文档](/zh/packages/pdf-preview/)。

## PPT 预览 —— `@marcusok/ppt-preview`

把 `.pptx` 文件拖入面板：解析在 Web Worker（Rust/WASM 引擎）内完成，幻灯片接近视口才懒渲染（远离视口的页自动释放 canvas 显存），首页出图与端到端耗时经 `onRendered` 上报。透明文本层让幻灯片可选中、可复制；页导航为 0-based（`goToSlide` / `nextSlide` / `prevSlide`），加密文件（MS-OFFCRYPTO Agile）在提供密码后自动解密。

<ClientOnly>
  <PackageDemo dir="ppt-preview" />
</ClientOnly>

完整用法见 [ppt-preview 文档](/zh/packages/ppt-preview/)。

## 进度遮罩 —— 内置于 `@marcusok/excel-exporter`

导出过程中看到的那个遮罩，这里经 `/overlay` 子路径直接驱动，任务与 Excel 无关。模拟任务前 1.5 秒不上报进度——那就是不确定态的旋转圆环；随后流式推进并切换为百分比条。两态渲染的是同一个实现、同一个毛玻璃面板，与 `overlay` 选项同源。

<ClientOnly>
  <OverlayDemo />
</ClientOnly>

完整用法见[进度遮罩指南](/zh/packages/excel-exporter/guide/11-overlay)。
