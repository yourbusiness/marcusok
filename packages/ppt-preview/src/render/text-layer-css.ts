/**
 * 页面布局与文本层的最小样式集 + 文档级幂等单例注入。
 *
 * 上游 buildPptxTextLayer 的 span/shape 容器定位全部内联 style.cssText
 * 写入（position/font/line-height/color:transparent 等，见 docs/ spike 记
 * 录 S4 段），本包只需提供：滚动布局（root/page）、文本层容器的覆盖定位
 * 与选中样式（user-select / ::selection 高亮）。上游文本层依赖内联 style
 * 属性——严格 CSP 的 style-src-attr 需放行 'unsafe-inline'（types.ts 的
 * styleNonce 注释有 CSP 完整口径；styleNonce 只覆盖本样式表）。
 *
 * 注入为 document 级幂等单例（id 查重）：多实例共享一份，不随 destroy
 * 删除（pdf-preview 同款取舍）。
 */

const STYLE_ID = "marcusok-ppt-preview-styles";

/** 滚动容器与页占位类名（消费方可用于自定义覆盖）。 */
export const ROOT_CLASS = "mop-ppt-root";
export const PAGE_CLASS = "mop-ppt-page";
/** 文本层容器类名（上游不写 class，定位样式由本包提供）。 */
export const TEXT_LAYER_CLASS = "mop-ppt-text-layer";

const CSS = `
.${ROOT_CLASS} {
  /* 连续滚动容器：幻灯片垂直堆叠、水平居中（deck 内所有页同尺寸） */
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 12px;
  padding: 12px 0;
}
.${PAGE_CLASS} {
  position: relative;
  direction: ltr;
  background-color: #fff;
  /* 页占位由 JS 按渲染宽度与 deck 比例显式定宽高（防滚动条跳动） */
}
.${PAGE_CLASS} > canvas {
  display: block;
}

/* ---- 文本层：上游 span 自带内联定位/透明色，这里只补容器与选中 ---- */
.${PAGE_CLASS} .${TEXT_LAYER_CLASS} {
  position: absolute;
  inset: 0;
  overflow: clip;
  z-index: 2;
  user-select: text;
  -webkit-user-select: text;
}
.${PAGE_CLASS} .${TEXT_LAYER_CLASS} span {
  cursor: text;
}
.${PAGE_CLASS} .${TEXT_LAYER_CLASS} ::selection {
  background: rgb(0 0 255 / 0.25);
}
`;

/** 注入幂等单例样式表（多实例共享，不随 destroy 删除）。 */
export function ensurePptStyles(nonce?: string): void {
  const doc = document;
  if (doc.getElementById(STYLE_ID)) return;
  const style = doc.createElement("style");
  style.id = STYLE_ID;
  if (nonce) style.setAttribute("nonce", nonce);
  style.textContent = CSS;
  doc.head.appendChild(style);
}
