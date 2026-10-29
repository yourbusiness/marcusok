/**
 * @marcusok/ppt-preview —— 浏览器端 .pptx 只读预览。
 *
 * 高层 API：`createPptPreview(container, { source, ... })` 一行接入，内部
 * 完成格式嗅探（sniff 前置分流）、source 防御性复制（引擎会 transfer 输入
 * buffer）、错误归一（上游错误码 → 5 个公开错误码）、连续滚动懒渲染
 * （IntersectionObserver，视口外页释放 canvas 显存）与文本层单趟收集
 * （正式渲染的 onTextRun 同趟构建，支持选中复制）。
 *
 * 渲染宽度在 create 时一次定型，容器尺寸变化不自动重排——重新 create 即可
 * （首版无 ResizeObserver，见 limits）。支持 OLE2/MS-OFFCRYPTO Agile 加密
 * 文件的 password 解密（主线程 WebCrypto，安全上下文要求见 limits）。
 */
export { createPptPreview } from "./render/render";
export type {
  PptPreviewOptions,
  PptPreviewInstance,
  PptPreviewError,
  PptPreviewErrorCode,
  PptRenderedInfo,
  PptRenderOptions,
  PptResourceLimits,
  PptSource,
} from "./types";
