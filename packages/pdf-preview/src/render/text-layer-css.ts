/**
 * 文本层与页面布局的最小样式集 + 文档级幂等单例注入。
 *
 * pdf.js 的 TextLayer **不自带定位样式**：span 的定位/字号以内联 style +
 * CSS 变量计算（`--font-height`、`--scale-x` 等由 JS 内联写入 span），但
 * `--total-scale-factor` / `--scale-round-x/y` 变量的定义、span 的绝对
 * 定位/透明色/选中样式都在官方 viewer 的 web/pdf_viewer.css 里——不注入
 * 本样式集，文本层会错位、字号为 0、选中无高亮（上游 pdf_viewer.css 的
 * .textLayer/.pdfViewer .page 段收窄，标注对齐上游 6.4.299）。
 *
 * 注入为 document 级幂等单例（id 查重）：多实例共享一份，不随 destroy
 * 删除。styleNonce 只覆盖本包注入的样式；pdf.js 的 FontLoader 运行时另
 * 行注入 @font-face 样式（adoptedStyleSheets 或 <head><style>），不在
 * 本 nonce 覆盖范围（types.ts 的 styleNonce 注释有 CSP 完整口径）。
 */

const STYLE_ID = "marcusok-pdf-preview-styles";

/** 页容器与布局类名（消费方可用于自定义覆盖）。 */
export const ROOT_CLASS = "mop-pdf-root";
export const PAGE_CLASS = "mop-pdf-page";

const CSS = `
.${ROOT_CLASS} {
  /* 连续滚动容器：页垂直堆叠、水平居中（PDF 允许逐页尺寸不同） */
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 12px;
  padding: 12px 0;
}
.${PAGE_CLASS} {
  /* 官方 viewer 的 .page 变量链：JS 侧按页写 --scale-factor（viewport.scale），
   * 文本层 span 的内联样式按 --total-scale-factor 计算字号与坐标 */
  --user-unit: 1;
  --total-scale-factor: calc(var(--scale-factor) * var(--user-unit));
  --scale-round-x: 1px;
  --scale-round-y: 1px;

  position: relative;
  direction: ltr;
  background-color: #fff;
}
.${PAGE_CLASS} > canvas {
  display: block;
}

/* ---- textLayer（官方 .textLayer 段收窄，去掉 highlight/annotation 相关） ---- */
.${PAGE_CLASS} .textLayer {
  position: absolute;
  inset: 0;
  overflow: clip;
  opacity: 1;
  z-index: 2;
  color-scheme: only light;
  text-align: initial;
  line-height: 1;
  letter-spacing: normal;
  word-spacing: normal;
  text-size-adjust: none;
  forced-color-adjust: none;
  transform-origin: 0 0;
  caret-color: CanvasText;

  --min-font-size: 1;
  --text-scale-factor: calc(var(--total-scale-factor) * var(--min-font-size));
  --min-font-size-inv: calc(1 / var(--min-font-size));
}
.${PAGE_CLASS} .textLayer :is(span, br) {
  color: transparent;
  position: absolute;
  white-space: pre;
  cursor: text;
  transform-origin: 0% 0%;
  user-select: text;
}
/* span 字号机制：JS 内联写 --font-height/--scale-x/--rotate，此处按变量计算
 * font-size 与 transform——缺这条文本不可见（字号 0） */
.${PAGE_CLASS} .textLayer > :not(.markedContent),
.${PAGE_CLASS} .textLayer .markedContent span:not(.markedContent) {
  z-index: 1;
  --font-height: 0;
  font-size: calc(var(--text-scale-factor) * var(--font-height));
  --scale-x: 1;
  --rotate: 0deg;
  transform: rotate(var(--rotate)) scaleX(var(--scale-x))
    scale(var(--min-font-size-inv));
}
.${PAGE_CLASS} .textLayer .markedContent {
  display: contents;
}
.${PAGE_CLASS} .textLayer span[role="img"] {
  user-select: none;
  cursor: default;
}
.${PAGE_CLASS} .textLayer ::selection {
  background: rgb(0 0 255 / 0.25);
  color: transparent;
}
`;

/** 注入幂等单例样式表（多实例共享，不随 destroy 删除）。 */
export function ensurePdfStyles(nonce?: string): void {
  const doc = document;
  if (doc.getElementById(STYLE_ID)) return;
  const style = doc.createElement("style");
  style.id = STYLE_ID;
  if (nonce) style.setAttribute("nonce", nonce);
  style.textContent = CSS;
  doc.head.appendChild(style);
}
