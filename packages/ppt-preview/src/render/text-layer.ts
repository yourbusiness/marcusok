/**
 * 文本层构建：上游 buildPptxTextLayer 的薄包装。
 *
 * runs 在正式渲染 `renderSlide(canvas, i, { onTextRun })` 的同一趟收集
 * （上游回调同步产出 PptxTextRunInfo[]），随后一次性构建 DOM——不调用
 * `collectSlideRuns()`（该 API 在 main 模式会另建 OffscreenCanvas 再渲染
 * 一次，双倍渲染成本，见 docs/ spike 记录）。
 *
 * 超链接有意 inert：不传 onHyperlinkClick 回调，上游只写 cursor:text、
 * 不绑 click（上游源码实证）——首版只承诺选择/复制，链接开放需另行设计
 * external scheme 白名单与内部跳转。
 */
import { buildPptxTextLayer } from "@silurus/ooxml/pptx";
import type { PptxTextRunInfo } from "@silurus/ooxml/pptx";
import { TEXT_LAYER_CLASS } from "./text-layer-css";

/**
 * 为一页构建文本层节点（挂在占位 el 内、覆盖 canvas）。
 * 返回创建的容器 div（由调用方纳入页资源清理）。
 */
export function attachTextLayer(
  pageEl: HTMLElement,
  runs: PptxTextRunInfo[],
  cssWidth: number,
  cssHeight: number,
): HTMLDivElement {
  const div = document.createElement("div");
  div.className = TEXT_LAYER_CLASS;
  // 上游按 cssWidth/cssHeight 把 run 坐标（EMU）换算为 CSS px
  buildPptxTextLayer(div, runs, cssWidth, cssHeight);
  pageEl.appendChild(div);
  return div;
}
