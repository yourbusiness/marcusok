# createPdfPreview

```ts
import { createPdfPreview } from "@marcusok/pdf-preview";

const preview = createPdfPreview(container, options);
// preview: { destroy(): void }
```

把连续滚动、懒渲染的 PDF 预览挂载进 `container`（容器只是挂载点——预览渲染进内部 wrapper，`destroy()` 同步摘除）。

## PdfPreviewOptions

| 选项                  | 类型                                        | 默认值        | 说明                                                                                                                                                                                                                      |
| --------------------- | ------------------------------------------- | ------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `source`              | `File \| Blob \| Uint8Array \| ArrayBuffer` | —             | PDF 字节。读取失败归 `UNKNOWN`                                                                                                                                                                                            |
| `scale`               | `number \| "fit-width" \| "fit-page"`       | `"fit-width"` | 创建时解析一次。数值必须为正、字符串只能是 `"fit-width"`/`"fit-page"`（否则同步抛 `TypeError`）。隐藏容器（0 宽）退回 `1` 并 `console.warn`                                                                               |
| `textLayer`           | `boolean`                                   | `true`        | canvas 之上可选中复制的透明文本层                                                                                                                                                                                         |
| `workerSrc`           | `string`                                    | 推导          | worker 脚本 URL。默认：bundler 自动 emit 的锚点字面量，src 直跑形态另有 `../dist/` 运行时兜底。可传绝对 URL 或根相对/相对路径；相对路径会先按页面基址解析，再推导 cmaps/字体目录（见[快速上手](../guide/01-quick-start)） |
| `cMapUrl`             | `string`                                    | 推导          | CMap 目录 URL（以 `/` 结尾）。bundler 场景 **CJK 文档必传**；缺失静默劣化                                                                                                                                                 |
| `standardFontDataUrl` | `string`                                    | 推导          | 标准字体目录 URL（以 `/` 结尾）；PDF 引用非内嵌标准字体时需要                                                                                                                                                             |
| `wasmUrl`             | `string`                                    | —             | pdf.js wasm 目录（不随包分发；缺它 JPX/JBIG2/ICC 降级——见[限制](../guide/02-limits)）                                                                                                                                     |
| `iccUrl`              | `string`                                    | —             | ICC 配置目录（不随包分发；缺它色彩管理降级）                                                                                                                                                                              |
| `renderOptions`       | `PdfRenderOptions`                          | —             | `getDocument` 参数白名单透传（见下）                                                                                                                                                                                      |
| `styleNonce`          | `string`                                    | —             | 本包注入 `<style>` 的 nonce。**不覆盖** pdf.js 自身的运行时字体样式——见[限制](../guide/02-limits)                                                                                                                         |
| `onRendered`          | `(info: PdfRenderedInfo) => void`           | —             | 首个成功渲染页绘制完成时触发一次（懒渲染管线没有"全部页完成"事件；若第一页渲染失败，由后续第一张成功页触发）                                                                                                              |
| `onError`             | `(error: PdfPreviewError) => void`          | —             | 全部失败路径。`destroy()` 触发的取消静默（不是错误）                                                                                                                                                                      |

## PdfRenderOptions

允许透传的 `getDocument` 参数白名单：

| 选项              | 类型      | 说明                            |
| ----------------- | --------- | ------------------------------- |
| `disableFontFace` | `boolean` | 不加载内嵌字体（调试/兜底辅助） |
| `useSystemFonts`  | `boolean` | 优先用系统字体替代标准字体文件  |

其余参数由本包接管：`data`（来自 `source`）、`cMapUrl`/`cMapPacked`/`standardFontDataUrl`/`wasmUrl`/`iccUrl`（资产推导 + 顶层选项）、`enableXfa`（固定 `false`）、verbosity（保持上游默认）、以及仅对 url 源有意义的 `disableAutoFetch`/`rangeChunkSize`（内存源下无意义，不透传）。`isEvalSupported` 在 pdf.js 6.x 已不存在。

## PdfRenderedInfo

| 字段          | 说明                                                                    |
| ------------- | ----------------------------------------------------------------------- |
| `pageCount`   | 文档总页数（`numPages`）                                                |
| `firstPageMs` | 首个绘制完成的页：canvas + 文本层                                       |
| `durationMs`  | 从 `createPdfPreview` 调用到该页绘制完成（源读取 + 嗅探 + 解析 + 渲染） |

## 错误码

| 错误码               | 何时                                                                        |
| -------------------- | --------------------------------------------------------------------------- |
| `UNSUPPORTED`        | 头部 1024 字节内无 `%PDF-` 魔数，或输入为空                                 |
| `PASSWORD_PROTECTED` | 加密 PDF（上游 `PasswordException`；message 区分需要密码与密码被拒）        |
| `CORRUPT`            | 魔数合法但解析/渲染失败（`InvalidPDFException` 与未预期异常；保留 `cause`） |
| `UNKNOWN`            | 管线之外——源读取失败、worker/资产 URL 不可达、资产 URL 配置错误             |

## PdfPreviewInstance

```ts
interface PdfPreviewInstance {
  destroy(): void;
}
```

`destroy()` 断开观察器、取消在途 `RenderTask` 与文本层流、清零已渲染 canvas、调用 `loadingTask.destroy()`（终止 worker）并摘除 DOM。幂等；之后回调静默。共享的样式单例有意保留（其他实例可能仍在使用）。
