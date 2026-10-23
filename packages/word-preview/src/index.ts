/**
 * @marcusok/word-preview —— 浏览器端 docx 只读预览。
 *
 * 高层 API：`createWordPreview(container, { source, ... })` 一行接入，
 * 内部完成格式嗅探（sniff 前置）、错误归一、安全加固（altChunk 默认关、
 * 超链接 scheme 白名单、styleNonce）与多实例样式隔离。
 *
 * 不设低层纯数据模型 API：docx-preview 仅 `renderAsync` 是稳定 API，
 * `WordDocument` 类型是 any stub（上游 d.ts 自注 `//stub`），低层 parse
 * 出的内部模型无法等价暴露（与 excel-preview 的分层差异，见 limits）。
 */
export { createWordPreview } from "./render/render";
export type {
  WordPreviewOptions,
  WordPreviewInstance,
  WordPreviewError,
  WordPreviewErrorCode,
  WordRenderedInfo,
  WordRenderOptions,
  WordSource,
} from "./types";
