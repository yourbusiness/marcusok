/**
 * @marcusok/pdf-preview —— 浏览器端 PDF 只读预览。
 *
 * 高层 API：`createPdfPreview(container, { source, ... })` 一行接入，内部
 * 完成格式嗅探（sniff 前置）、错误归一、资产 URL 推导（worker/cmaps/
 * standard_fonts 三级 fallback）、连续滚动懒渲染（IntersectionObserver，
 * 视口外页释放 canvas 显存）与文本层样式自足（可选中复制）。
 *
 * 缩放为 create 时一次解析（fit-width/fit-page/数值），容器尺寸变化不自动
 * 重排——重新 create 即可（首版无 ResizeObserver，见 limits）。
 */
export { createPdfPreview } from "./render/render";
export type {
  PdfPreviewOptions,
  PdfPreviewInstance,
  PdfPreviewError,
  PdfPreviewErrorCode,
  PdfRenderedInfo,
  PdfRenderOptions,
  PdfScale,
  PdfSource,
} from "./types";
