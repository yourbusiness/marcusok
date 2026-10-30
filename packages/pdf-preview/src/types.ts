/**
 * @marcusok/pdf-preview 公共类型。
 *
 * 公开类型全部本包自定义（不 re-export pdfjs-dist 的类型），
 * `dist/index.d.ts` 自包含、不出现 `from "pdfjs-dist"` 外部引用——
 * npm 消费方零传递依赖（pdfjs-dist 是 devDependency、构建期整体打进
 * dist，word-preview 已验证的同款决策）。
 */

/** 输入源：文件对象或原始字节。 */
export type PdfSource = File | Blob | Uint8Array | ArrayBuffer;

/**
 * 缩放模式：
 * - 数值 —— 直接作为 pdf.js viewport scale（1 = 72dpi 原始尺寸，约 CSS
 *   96dpi 下偏小；常用 1.25-1.5）；
 * - `"fit-width"`（默认）—— 按容器宽与首页宽之比缩放，各页按自身尺寸
 *   等比渲染并水平居中（PDF 允许逐页尺寸不同）；
 * - `"fit-page"` —— 取容器宽/高比中较小者，整页落入容器。
 */
export type PdfScale = number | "fit-width" | "fit-page";

/**
 * `renderOptions` 白名单：pdf.js `getDocument` 参数中可透传的子集。
 * 以下参数由本包接管，传入无效：
 * - `data` —— 由顶层 `source` 转换；
 * - `cMapUrl` / `cMapPacked` / `standardFontDataUrl` / `wasmUrl` / `iccUrl`
 *   —— 由顶层同名选项 + 资产 URL 推导接管（见 assets.ts）；
 * - `enableXfa` —— 固定 `false`（XFA 表单首版不支持）；
 * - `isEvalSupported` —— pdf.js 6.x 已移除该参数（内嵌脚本改由 sandbox
 *   模块执行，本包不加载该模块，PDF 内嵌 JavaScript 永不执行）；
 * - `disableAutoFetch` / `rangeChunkSize` —— 仅对 url 源有意义，首版
 *   `source` 只有内存字节，不透传；
 * - `verbosity` —— 保留上游默认，本包不吞日志。
 */
export interface PdfRenderOptions {
  /** 不加载/应用文档内嵌字体（默认 false）。 */
  disableFontFace?: boolean;
  /** 优先使用系统字体替代标准字体文件（默认随上游）。 */
  useSystemFonts?: boolean;
}

export interface PdfPreviewOptions {
  source: PdfSource;
  /** 缩放模式（默认 `"fit-width"`，见 {@link PdfScale}）。 */
  scale?: PdfScale;
  /**
   * 是否渲染文本层（默认 true）：透明绝对定位文本与 canvas 像素对齐，
   * 支持选中复制；false 时纯 canvas。
   */
  textLayer?: boolean;
  /**
   * worker 脚本 URL 显式覆盖。默认按三级推导（见 limits 文档）：
   * 显式选项 → `new URL("./pdf.worker.min.mjs", import.meta.url)` 锚点
   * （bundler 会把它 emit 为资产；裸引用时指向包内 dist）→ src 直跑
   * 场景的 `../dist/` 运行时兜底。bundler 场景通常零配置。
   *
   * 可传绝对 URL 或根相对/相对路径；相对路径会按页面 `document.baseURI`
   * 解析后再推导 cmaps/standard_fonts 目录。
   */
  workerSrc?: string;
  /**
   * cmaps 目录 URL（以 `/` 结尾）。中文/CJK PDF 的硬需求——缺失时 CID
   * 字体文本静默劣化（空白/乱码），不报错。默认从 worker URL 推导
   * （bundler 资产化场景下推导路径可能不含 cmaps，需显式传入或用 copy
   * 插件同步资产，见 limits 的场景矩阵）。
   */
  cMapUrl?: string;
  /**
   * standard_fonts 目录 URL（以 `/` 结尾）。引用非内嵌标准字体
   * （Helvetica 等 14 种）的 PDF 需要；默认从 worker URL 推导，同
   * {@link cMapUrl} 的 bundler 场景注意事项。
   */
  standardFontDataUrl?: string;
  /**
   * wasm 目录 URL（以 `/` 结尾）。本包不分发 wasm 资产（JPX/JBIG2 图像
   * 解码、ICC 色彩管理）；不传时上游走 JS 慢速兜底或降级（qcms 缺失仅
   * 警告）。需要完整保真时自托管 pdfjs-dist 的 wasm/ 目录并传入。
   */
  wasmUrl?: string;
  /**
   * ICC 色彩配置目录 URL（以 `/` 结尾）。与 wasmUrl 同属不分发资产，
   * 缺失时色彩管理降级（上游仅警告）。
   */
  iccUrl?: string;
  /**
   * getDocument 参数白名单子集（见 {@link PdfRenderOptions}）。
   */
  renderOptions?: PdfRenderOptions;
  /**
   * CSP nonce：本包注入的文本层/布局 `<style>` 会带上 nonce。注意 pdf.js
   * 的 FontLoader 运行时会另行向 adoptedStyleSheets 或 `<head>` 注入
   * @font-face 样式，不受本 nonce 覆盖——严格 CSP 需放行 `style-src`；
   * 文本层 span 依赖内联 style 属性，需 `style-src-attr 'unsafe-inline'`；
   * 内嵌字体走 data: URL，需 `font-src data:`（见 limits 文档）。
   */
  styleNonce?: string;
  /**
   * 首个成功渲染页完成回调（懒渲染形态下没有"全部页完成"事件——视口外
   * 的页尚未渲染，这是有意设计）。注意是"首个**成功**页"：若第一页渲染
   * 失败（已报 onError、实例仍可继续），后续滚到的第一张成功页会触发本
   * 回调，其 `firstPageMs` 计的是那张页的渲染耗时。
   * - `pageCount` = 文档总页数（`numPages` 直出）；
   * - `firstPageMs` = 该页 canvas + 文本层渲染耗时（用户感知的出图时间）；
   * - `durationMs` = 从 createPdfPreview 调用到该页完成的端到端耗时
   *   （含源读取、嗅探、getDocument、渲染）。
   */
  onRendered?: (info: PdfRenderedInfo) => void;
  /** 错误回调。destroy 触发的取消静默，不报错。 */
  onError?: (error: PdfPreviewError) => void;
}

export interface PdfRenderedInfo {
  /** 文档总页数。 */
  pageCount: number;
  /** 触发回调的那页 canvas + 文本层渲染耗时（毫秒）。 */
  firstPageMs: number;
  /** 从 createPdfPreview 调用到该页渲染完成的端到端耗时（毫秒）。 */
  durationMs: number;
}

export type PdfPreviewErrorCode =
  /** 头部 1024 字节窗口内无 `%PDF-` 魔数（不是 PDF，或垃圾前缀超窗）。 */
  | "UNSUPPORTED"
  /** 加密 PDF：首版仅识别并报错，不提示输入密码、不解密。 */
  | "PASSWORD_PROTECTED"
  /** 魔数合法但解析失败（结构损坏、InvalidPDFException 及未预期异常）。 */
  | "CORRUPT"
  /** 渲染管线之外的失败（源读取失败、worker/资产不可达、资产 URL 配置错误）。 */
  | "UNKNOWN";

export interface PdfPreviewError {
  code: PdfPreviewErrorCode;
  message: string;
  /** 原始错误（诊断用）。 */
  cause?: unknown;
}

export interface PdfPreviewInstance {
  /**
   * 销毁实例：断开懒渲染观察、取消在途渲染任务与文本层流、释放 canvas
   * 位图、终止 pdf.js worker（loadingTask.destroy）并摘除容器内本包 DOM。
   * 销毁后所有回调静默（取消不是错误，不触发 onError）。幂等。
   *
   * 已注入的文本层样式表是文档级幂等单例（多实例共享），不随 destroy
   * 删除。容器尺寸变化不会自动重排（destroy 后重新 create 即可）。
   */
  destroy(): void;
}
