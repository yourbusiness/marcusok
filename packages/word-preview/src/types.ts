/**
 * @marcusok/word-preview 公共类型。
 *
 * 公开类型全部本包自定义（不 re-export docx-preview 的 `Options`），
 * `dist/index.d.ts` 自包含、不出现 `from "docx-preview"` 外部引用——
 * npm 消费方零传递依赖（docx-preview 与 jszip 均为 devDependencies、
 * 构建期打进 dist）。
 */

/** 输入源：文件对象或原始字节。 */
export type WordSource = File | Blob | Uint8Array | ArrayBuffer;

/**
 * `renderOptions` 白名单：docx-preview 上游 `Options`（20 项）中可透传的
 * 13 项子集（含 `ignoreLastRenderedPageBreak`——本包改了它的默认值，但
 * 允许调用方改回）。其余 7 项由本包硬接管，不可透传：
 * - `className` —— 本包为每个实例生成唯一类名（同页多实例的全局 CSS 串扰
 *   隔离，见 createWordPreview）；
 * - `useBase64URL` —— 固定 `true`（上游 0 处 revokeObjectURL，blob URL 会
 *   泄漏；base64 的内存代价约 +33% 见 limits 文档）；
 * - `renderAltChunks` —— 由 `allowAltChunks` 选项接管（默认关闭，XSS 加固）；
 * - `h` / `experimental` / `debug` / `trimXmlDeclaration` —— experimental
 *   hook、实验特性、日志开关与内部细节，不进公开面。
 */
export interface WordRenderOptions {
  /** 按分页符断页（默认 true）。false 时整篇渲染为单一 section——
   * `onRendered` 的 `pageCount` 恒为 1。 */
  breakPages?: boolean;
  /** 渲染进 `.docx-N-wrapper` 页面容器（默认 true）。false 改变包裹结构。 */
  inWrapper?: boolean;
  hideWrapperOnPrint?: boolean;
  /** 忽略页面宽度（默认 false）。 */
  ignoreWidth?: boolean;
  /** 忽略页面高度（默认 false）。 */
  ignoreHeight?: boolean;
  /** 不加载/应用文档内嵌字体（默认 false）。 */
  ignoreFonts?: boolean;
  /** 忽略 `lastRenderedPageBreak` 标记（上游默认 true；本包默认 false，
   * 贴近 Word 原版分页——该标记由编辑器存盘写入、非排版计算，与真实
   * 分页可能有出入）。 */
  ignoreLastRenderedPageBreak?: boolean;
  renderHeaders?: boolean;
  renderFooters?: boolean;
  renderFootnotes?: boolean;
  renderEndnotes?: boolean;
  renderComments?: boolean;
  /** 渲染修订（track changes，默认 false）。 */
  renderChanges?: boolean;
}

export interface WordPreviewOptions {
  source: WordSource;
  /**
   * docx-preview 选项子集（白名单见 {@link WordRenderOptions}）。
   * `className` / `useBase64URL` / `renderAltChunks` 由本包接管，
   * 传入无效。
   */
  renderOptions?: WordRenderOptions;
  /**
   * CSP nonce：渲染注入的 `<style>` 会以"换节点"方式重建并带上 nonce
   * （CSP 对 `<style>` 的拦截发生在节点插入时，事后补属性不会生效）。
   * 注意 docx-preview 大量使用内联 style 属性，严格 CSP 下仍需放行
   * `style-src-attr 'unsafe-inline'`；`useBase64URL: true` 的图片/字体
   * 需要 `img-src data:` / `font-src data:`（见 limits 文档）。
   */
  styleNonce?: string;
  /**
   * 是否渲染 altChunk（文档内嵌 HTML 片段，经 `<iframe srcdoc>` 呈现）。
   * 默认 **false**：上游的 iframe 无 `sandbox` 属性，不可信文档可借
   * `srcdoc` 执行脚本。开启仅限受信文档——本包会在渲染完成后补
   * `sandbox` 属性，但属性赋值与首帧加载的先后不保证（各浏览器行为
   * 不一，残余时序风险见 limits 文档）；iframe 内部文档也不在超链接
   * scheme 白名单的覆盖范围内。
   */
  allowAltChunks?: boolean;
  /**
   * 渲染完成回调。`durationMs` 从 `createWordPreview` 调用起计——含源
   * 读取与格式嗅探，再加**含图片/内嵌字体加载**的渲染本身（上游
   * resolve 前 await 全部资源任务），非纯 DOM 构建。
   * `pageCount` 按容器内 `section` 元素计数——仅在默认分页选项
   * （`breakPages: true` / `inWrapper: true`）下有意义。
   */
  onRendered?: (info: WordRenderedInfo) => void;
  /** 错误回调（格式/解析/渲染失败）。destroy 触发的取消静默，不报错。 */
  onError?: (error: WordPreviewError) => void;
}

export interface WordRenderedInfo {
  /** 渲染出的 section（页）元素数；仅默认分页选项下有意义。 */
  pageCount: number;
  /** 从 createWordPreview 调用到渲染完成的耗时（毫秒，含源读取、
   * 嗅探与图片/字体等资源加载）。 */
  durationMs: number;
}

export type WordPreviewErrorCode =
  /** 非 zip 容器，或 zip 内缺少 Word 必需部件（不是 Word 文档）。 */
  | "UNSUPPORTED"
  /** OLE2 复合文档：.doc 老格式与标准加密（Agile AES）的 docx 同为该魔数，
   * 两者均不支持（合并表述）。 */
  | "LEGACY_FORMAT"
  /** zip 头 general purpose bit 0 置位（直接 zip 加密形态）。 */
  | "PASSWORD_PROTECTED"
  /** zip 结构损坏、必需部件存在但解析/渲染失败、上游未预期异常。 */
  | "CORRUPT"
  /** 渲染管线之外的意外失败（如输入源 Blob 读取失败）。 */
  | "UNKNOWN";

export interface WordPreviewError {
  code: WordPreviewErrorCode;
  message: string;
  /** 原始错误（诊断用）。 */
  cause?: unknown;
}

export interface WordPreviewInstance {
  /**
   * 卸载渲染 DOM。上游无取消机制（0 处 abort），destroy 无法中断在途的
   * 同步渲染段，但本包的渲染写入内部私有 wrapper——destroy 同步摘除它，
   * 迟到的渲染结果只会写入已脱离文档的节点，随后被丢弃；destroy 后所有
   * 回调静默。取消不是错误，不会触发 onError。
   */
  destroy(): void;
}
