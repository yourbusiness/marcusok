/**
 * @marcusok/ppt-preview 公共类型。
 *
 * 公开类型全部本包自定义（不 re-export @silurus/ooxml 的类型），
 * `dist/index.d.ts` 自包含、不出现 `from "@silurus/ooxml"` 外部引用——
 * npm 消费方零传递依赖（上游是 devDependency、构建期整体打进 dist，
 * pdf/word-preview 已验证的同款决策）。上游 0.x 发布频繁，钉死公开
 * 类型同时把升级影响隔离在本包内部。
 */

/** 输入源：文件对象或原始字节。 */
export type PptSource = File | Blob | Uint8Array | ArrayBuffer;

/**
 * `renderOptions` 白名单：上游 `PptxPresentation.load` LoadOptions 中可
 * 透传的子集。以下参数由本包接管，传入无效：
 * - 数据源 —— 由顶层 `source` 归一化（上游签名只收 `string | ArrayBuffer`）；
 * - `mode` —— 固定 `"main"`（解析仍在 module Web Worker 中、渲染在主线程）；
 * - `progressiveLayout` —— 首版不启用（slideCount 与占位数保持稳定）；
 * - `password` / `useGoogleFonts` / `wasmUrl` —— 顶层同名选项接管；
 * - `math` / `threeD` / `regionMap` / `chartEx` / `tiff` —— 可选渲染器，
 *   首版不注入（未注入时对应内容走上游降级，见 limits）；
 * - `maxZipEntryBytes` —— 上游已废弃，用 `resourceLimits.maxArchiveEntryBytes`。
 */
export interface PptRenderOptions {
  /**
   * CJK 缺字回退目标字形（默认 `"auto"`，按内容自动判定）。中文字体
   * 未内嵌时的兜底渲染，`"sc"` 简体 / `"tc"` 繁体台湾 / `"hk"` 繁体香港 /
   * `"jp"` 日文 / `"kr"` 韩文。
   */
  cjkFallback?: "auto" | "kr" | "sc" | "tc" | "hk" | "jp";
  /**
   * 压缩包资源上限（合法文件超预算时按 `UNSUPPORTED` 上报，不是文件
   * 损坏）。字段与上游现行 `OoxmlResourceLimits` 对齐；`null` 表示显式
   * 关闭某项上限。
   */
  resourceLimits?: PptResourceLimits;
  /** 解析 Worker 单次响应超时（毫秒），同时约束正常加载与加载期间 destroy 的最坏等待时间。 */
  workerTimeoutMs?: number;
}

/** 压缩包资源上限（透传上游，见 {@link PptRenderOptions.resourceLimits}）。 */
export interface PptResourceLimits {
  /** 单个 zip 条目解压后最大字节数。 */
  maxArchiveEntryBytes?: number | null;
  /** 全部条目解压总量上限。 */
  maxTotalInflatedBytes?: number | null;
  /** zip 条目数量上限。 */
  maxArchiveEntries?: number | null;
}

export interface PptPreviewOptions {
  source: PptSource;
  /**
   * 打开密码（仅支持 OLE2/MS-OFFCRYPTO Agile 加密的 .pptx）。密钥派生
   * （SHA-512 迭代）与 AES-CBC 解密在主线程执行，通常增加约一秒解析
   * 前耗时；需要安全上下文（HTTPS/localhost 的 WebCrypto）。ZIP 条目
   * 加密不在承诺范围内。
   */
  password?: string;
  /**
   * 幻灯片渲染宽度（CSS 像素，默认 960），小数向下取整。高度按 deck 实际
   * 宽高比等比推导；必须是有限正数，非法值同步抛 `TypeError`。deck 尺寸
   * 异常导致比例无效时按内容错误上报，不会生成 NaN 占位。
   */
  slideWidth?: number;
  /**
   * devicePixelRatio 覆盖（默认取 `window.devicePixelRatio`）。必须是
   * 有限正数，非法值同步抛 `TypeError`。
   */
  dpr?: number;
  /**
   * 是否渲染文本层（默认 true）：透明绝对定位文本与 canvas 像素对齐，
   * 支持选中复制；false 时纯 canvas。文本 runs 在正式渲染的同一趟
   * `renderSlide` 中收集（`onTextRun`），不做二次渲染。
   */
  enableTextSelection?: boolean;
  /**
   * 是否允许加载 Google Fonts 网络字体（默认 false，不发起任何外部
   * 请求）。开启后缺失字体可能从 fonts.googleapis.com 拉取替代字形，
   * CSP 需放行相应域名，隐私影响见 limits 文档。
   */
  useGoogleFonts?: boolean;
  /**
   * WASM 资产 URL 显式逃生门。不传时按 `new URL("./pptx_parser_bg.wasm",
   * import.meta.url)` 锚点定位（bundler 会 emit 包内资产；裸引用时指向
   * 包 dist）。仅在自托管 wasm 或资产路径被改写时需要传入。
   */
  wasmUrl?: string | URL;
  /**
   * CSP nonce：本包注入的布局/文本层 `<style>` 会带上 nonce（受
   * `style-src` 约束）。上游文本层与页面占位通过 CSSOM（`element.style`）
   * 写定位，Chromium 下不受 `style-src-attr` 拦截；只有以 HTML `style`
   * 属性形式存在的内联样式才需要 `style-src-attr 'unsafe-inline'`。其他
   * 内核若对 CSSOM 写入也施加该指令，可放行 unsafe-inline 或关闭文本层，
   * 见 limits 文档。
   *
   * 多实例边界：样式表是文档级幂等单例（首个实例注入、不随 destroy 删除），
   * nonce 先到先得——首个实例的值生效，后续实例传入的 nonce 不会被应用。
   */
  styleNonce?: string;
  /** 解析加载白名单子集（见 {@link PptRenderOptions}）。 */
  renderOptions?: PptRenderOptions;
  /**
   * 首张幻灯片渲染完成回调（懒渲染形态下没有"全部页完成"事件——视口
   * 外的页尚未渲染，这是有意设计）。
   * - `slideCount` = deck 原始总页数（含隐藏页）；
   * - `firstSlideMs` = 首张成功渲染页的 canvas + 文本层耗时；
   * - `durationMs` = 从 createPptPreview 起的端到端耗时（含源读取、嗅探、
   *   解密、**全 deck preflight** 与渲染——非渐进加载会预检全部页，
   *   懒渲染只延迟绘制，不等于懒解析）。
   *
   * 每实例最多触发一次，由首张**成功**页触发；全部页失败、空 deck
   * （slideCount=0），或容器在创建时不可见/尺寸为 0（例如 `display:none`
   * 的 tab 面板，占位永远进不了渲染带）都不触发。等容器可见后
   * IntersectionObserver 才会开始渲染并触发回调。
   */
  onRendered?: (info: PptRenderedInfo) => void;
  /** 错误回调。destroy 触发的取消静默，不报错；destroy 后回调静默。 */
  onError?: (error: PptPreviewError) => void;
}

export interface PptRenderedInfo {
  /** deck 原始总页数（含隐藏页）。 */
  slideCount: number;
  /** 首张成功渲染页的 canvas + 文本层耗时（毫秒）。 */
  firstSlideMs: number;
  /** 从 createPptPreview 调用到首张成功页渲染完成的端到端耗时（毫秒）。 */
  durationMs: number;
}

export type PptPreviewErrorCode =
  /** 非 PPTX 格式、老式 .ppt 之外的无法识别容器，以及密码正确也无法解决的不支持项（含资源超限）。 */
  | "UNSUPPORTED"
  /** 老式二进制 .ppt（引擎确认为 legacy 后使用；OLE2 输入不提前归到此码）。 */
  | "LEGACY_FORMAT"
  /** 加密 .pptx：缺密码或密码错误（message 区分两种情况；仅承诺 Agile 路径）。 */
  | "PASSWORD_PROTECTED"
  /** 魔数合法但结构损坏（可稳定区分的内容损坏才归此码，其余归 UNKNOWN）。 */
  | "CORRUPT"
  /** 渲染管线之外的失败：源读取、Worker、WASM、CSP、资产 URL、未知异常。 */
  | "UNKNOWN";

export interface PptPreviewError {
  code: PptPreviewErrorCode;
  message: string;
  /** 原始错误（诊断用）。 */
  cause?: unknown;
}

export interface PptPreviewInstance {
  /**
   * 销毁实例：断开懒渲染观察、标记在途渲染过期并清理、清 canvas/文本层
   * DOM、销毁引擎（终止解析 Worker）、摘除容器内本包 DOM。幂等；销毁后
   * 所有回调静默。
   *
   * 注意加载期间的 destroy 不是立即取消（上游 load 无 AbortSignal）：
   * 旧结果作废，load 迟到 resolve 后立即销毁引擎实例；`workerTimeoutMs`
   * 影响最坏等待。已注入的文本层样式表是文档级幂等单例（多实例共享），
   * 不随 destroy 删除。容器尺寸变化不自动重排（重建实例即可）。
   */
  destroy(): void;
  /**
   * 滚动到指定页（0-based，与 `slideCount` 同口径、含隐藏页、保持原始
   * 页序）。识别到滚动容器时只滚动该容器；容器本身不滚动（普通文档流）
   * 时回退滚动视口（`window`），此时外层页面会联动。
   *
   * 混合边界（有意设计，注意）：
   * - 非整数参数 → 同步抛 `TypeError`（校验先于状态检查，destroy 后同样抛）；
   * - 加载完成前或 destroy 后 → no-op（此时页占位尚不存在）；
   * - 加载完成后越界（`index < 0 || index >= slideCount`）→ 抛 `RangeError`。
   */
  goToSlide(index: number): void;
  /** 滚到下一页（基于当前最可见页，而非最后一次 goToSlide；边界/加载前/destroy 后为 no-op）。 */
  nextSlide(): void;
  /** 滚到上一页（口径同 {@link nextSlide}）。 */
  prevSlide(): void;
  /**
   * deck 原始总页数（含隐藏页）。加载完成前及 destroy 后为 0；空 deck
   * 加载完成后为 0。
   */
  get slideCount(): number;
}
