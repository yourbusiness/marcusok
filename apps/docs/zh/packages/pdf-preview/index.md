# @marcusok/pdf-preview

浏览器端 PDF 只读预览——连续滚动 + 逐页懒渲染、可选中复制的文本层、归一化错误码。渲染引擎为 [pdf.js](https://github.com/mozilla/pdf.js)（pdfjs-dist 6.4.299，构建期打进本包 `dist`，消费方零传递依赖），本包在其上叠加格式嗅探、资产 URL 推导与视口驱动的显存管理。

## 能力

| 能力        | 说明                                                                                                                                                                  |
| ----------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 格式        | 任何符合规范的 PDF：`%PDF-` 魔数在头部 1024 字节内搜索（Adobe 自家阅读器的窗口口径，允许文件头前有垃圾字节）。魔数超出窗口的文件归 `UNSUPPORTED`（有意的输入收紧）    |
| 懒渲染      | 连续滚动；每页占位尺寸按**该页自身**的 viewport 计算（混合尺寸文档不跳动）。页面接近视口（1 屏预载带）即渲染，远离视口（超 3 屏）即释放 canvas 显存、文本层流与页资源 |
| 文本层      | 与 canvas 像素对齐的透明文本层（pdf.js `TextLayer`）——像真实文本一样选中、复制、检索。`textLayer: false` 一行关闭                                                     |
| 缩放        | `scale` 支持数值 / `"fit-width"`（默认）/ `"fit-page"`，创建时按容器与首页一次解析；异尺寸页面保持等比居中                                                            |
| Worker 解析 | 解析与渲染调度在 pdf.js Web Worker 内完成；主线程只承担 canvas 绘制与 DOM 更新                                                                                        |
| 安全加固    | 引擎运行前先嗅探；加密 PDF 归 `PASSWORD_PROTECTED`（首版不解密）；不加载 pdf.js sandbox 模块，内嵌 PDF JavaScript 永不执行；注释层（超链接/表单）首版不渲染           |
| 取消        | `destroy()` 断开观察器、取消在途渲染任务与文本层流、清零已渲染 canvas，并终止 worker（`loadingTask.destroy()`）；随后回调静默（取消不是错误，不触发 `onError`）       |
| 错误        | 携错误码：`UNSUPPORTED`、`PASSWORD_PROTECTED`、`CORRUPT`、`UNKNOWN`                                                                                                   |

## 安装

```bash
pnpm add @marcusok/pdf-preview
```

pdf.js 引擎、worker 与 CJK/标准字体资产目录全部随包 `dist` 分发（载荷约 4.5 MB，另含 sourcemap；`cmaps` 是中文/CJK PDF 的硬需求）。Vite/webpack 下 worker **零配置**；`cmaps`/`standard_fonts` 目录在 bundler 场景需显式传 URL 或配 copy 插件——见[快速上手](/zh/packages/pdf-preview/guide/01-quick-start)。

## 最小示例

```ts
import { createPdfPreview } from "@marcusok/pdf-preview";

const preview = createPdfPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  scale: "fit-width", // 默认；也接受数值与 "fit-page"
  onRendered: ({ pageCount, firstPageMs, durationMs }) =>
    console.log(pageCount, firstPageMs, durationMs),
  onError: (e) => console.error(e.code, e.message),
});

preview.destroy(); // 取消在途任务 + 终止 worker
```

有意**不提供低层数据模型 API**：pdf.js 的文档模型是绑定在活跃 worker 上的 promise 代理，无法诚实地抽成纯数据的 `parse()` 结果。这一点与提供 `parseWorkbookBytes` 的 [@marcusok/excel-preview](/zh/packages/excel-preview/) 不同，该模式见其文档。

<ClientOnly>
<PdfPreviewDemo />
</ClientOnly>

## 首版不做

- **解密**——加密 PDF 会被识别并上报（`PASSWORD_PROTECTED`）；无密码输入、不解密。
- **注释层**——超链接、表单控件与批注不渲染（收窄攻击面与首版范围；列为演进项）。
- **容器重排**——缩放在创建时一次解析；容器尺寸变化不自动重拟合（destroy 后重建即可）。
- **JPX/JBIG2 图像与 ICC 色彩管理**——pdf.js 的 wasm 资产有意不分发；受影响文档降级（JS 兜底解码/无色彩管理），自托管并传 `wasmUrl`/`iccUrl` 可恢复完整保真。见[限制](/zh/packages/pdf-preview/guide/02-limits)。
