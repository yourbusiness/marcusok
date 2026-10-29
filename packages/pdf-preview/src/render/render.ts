/**
 * createPdfPreview 的实现：对 pdf.js 的薄封装 + 连续滚动懒渲染。
 *
 * 封装职责（偏离上游默认/补足缺口的部分，均见包 README / limits）：
 * 1. sniff 前置（../sniff.ts）——上游对非 PDF 输入的报错模糊且发生在
 *    worker 内部，先归一为 UNSUPPORTED；
 * 2. 资产 URL 推导（../assets.ts）——worker/cmaps/standard_fonts 三级
 *    fallback，bundler 场景 worker 零配置；
 * 3. 懒渲染状态机——每页占位 div（逐页 viewport 尺寸，防滚动条跳动）+
 *    双 IntersectionObserver（1 屏预载渲染 / 3 屏外释放 canvas 显存）+
 *    串行渲染队列（scroll 洪峰下同时最多 1 页在渲，离带条目出队即弃）；
 * 4. 文本层样式自足（text-layer-css.ts）——上游 TextLayer 不带定位样式，
 *    官方 viewer CSS 的 .textLayer 段必须外部提供；
 * 5. 错误归一（normalizeError）——PasswordException/InvalidPDFException
 *    映射为公开错误码，worker/资产加载失败归 UNKNOWN；
 * 6. destroy 全链路取消——上游有真实取消机制（RenderTask.cancel +
 *    TextLayer.cancel + loadingTask.destroy），destroy 后回调静默。
 */
import {
  GlobalWorkerOptions,
  InvalidPDFException,
  PasswordException,
  RenderingCancelledException,
  TextLayer,
  getDocument,
} from "pdfjs-dist";
import type {
  PDFDocumentLoadingTask,
  PDFDocumentProxy,
  PDFPageProxy,
  RenderTask,
} from "pdfjs-dist";
import { resolveAssetUrls, type PdfAssetUrls } from "../assets";
import { sniffPdf } from "../sniff";
import type {
  PdfPreviewInstance,
  PdfPreviewOptions,
  PdfPreviewError,
  PdfSource,
} from "../types";
import { PAGE_CLASS, ROOT_CLASS, ensurePdfStyles } from "./text-layer-css";

/** 渲染带（进入即入队渲染）：视口上下各预载 1 屏。 */
const RENDER_MARGIN = "100% 0px 100% 0px";
/** 保留带（离开才释放 canvas）：视口上下各 3 屏——与渲染带之间的 1-3 屏
 * 缓冲区避免快速滚动时的渲染/释放抖动。 */
const RELEASE_MARGIN = "300% 0px 300% 0px";

type PageState = "empty" | "queued" | "rendering" | "rendered";

/**
 * 判断滚动祖先是否真的可能裁剪内容，而不是"随内容一起长高"的 overflow:auto
 * 包装层（典型如 Ant Design 的 ant-card-body）。
 *
 * 只有 overflow:auto 不够：内容驱动的包装层计算样式同样是 overflow:auto，
 * 但它永远不滚动；把它选作 IntersectionObserver 的 root 后，root 的矩形
 * 就是整份文档的高度，所有页都判定为"在视口内"——懒渲染退化为全量渲染
 * （play 的 200 页 demo 实测 200/200 个 canvas）。真滚动的祖先至少有下面
 * 之一：受 max-height 约束、内联定高（非百分比）、当前已产生溢出滚动。
 * 都不满足时继续向上找，最终回退视口（root:null）。
 */
function hasDefiniteHeight(
  el: HTMLElement,
  style: CSSStyleDeclaration,
): boolean {
  const inlineHeight = el.style.height;
  const inlineMaxHeight = el.style.maxHeight;
  return Boolean(
    (style.maxHeight && style.maxHeight !== "none") ||
    (inlineMaxHeight && inlineMaxHeight !== "none") ||
    // 百分比高度在父级高度 auto 时并不定高，不能作为"会滚动"的证据
    (inlineHeight && inlineHeight !== "auto" && !inlineHeight.endsWith("%")) ||
    (el.clientHeight > 0 && el.scrollHeight > el.clientHeight),
  );
}

/**
 * 探测容器最近的、真正会裁剪内容的滚动祖先（overflow 为 auto/scroll）。
 *
 * IntersectionObserver 的 root 必须用它而不是默认视口：root:null 下
 * 交叉矩形会被滚动祖先裁剪，rootMargin 的预载带随之塌缩为该祖先的
 * 可视框——固定高度 overflow:auto 预览面板（最常见形态）里等于零预载，
 * 滚动时页面弹入（浏览器验收实测发现）。以滚动容器为 root 后
 * rootMargin 按容器高度展开，1 屏预载/3 屏释放才真正生效。
 *
 * 只接受 hasDefiniteHeight 的祖先（见其注释）——否则内容驱动的
 * overflow:auto 包装层会让懒渲染退化为全量渲染。
 */
function findScrollRoot(start: HTMLElement): HTMLElement | null {
  let node: HTMLElement | null = start;
  while (node && node !== document.documentElement) {
    const style = window.getComputedStyle(node);
    const scrollable =
      style.overflowY === "auto" ||
      style.overflowY === "scroll" ||
      style.overflow === "auto" ||
      style.overflow === "scroll";
    if (scrollable && hasDefiniteHeight(node, style)) return node;
    node = node.parentElement;
  }
  return null;
}

interface PageEntry {
  /** 1-based 页码。 */
  num: number;
  el: HTMLDivElement;
  page: PDFPageProxy | null;
  state: PageState;
  /** 当前在渲染带内（renderIO 判定）——出队时复核，离带即弃。 */
  wanted: boolean;
  /** 释放请求到达时仍在渲染：渲染完成后立即释放。 */
  pendingRelease: boolean;
  canvas: HTMLCanvasElement | null;
  textDiv: HTMLDivElement | null;
  /** 在途文本层（成功渲染后置空）；release/destroy 必须 cancel 掉。 */
  textLayer: TextLayer | null;
  task: RenderTask | null;
}

// pdf.js 的 GetDocRequest 以 [data.buffer] 为 transfer list（pdf.mjs 源码
// 实证），transfer 后调用方持有的字节立即 detach（length=0）。字节型
// source 必须复制一份再交给引擎——否则同一 source 二次 create（本包文档
// 倡导的"重 create 即可"路径）会在 sniff 阶段读到空字节、以
// "File is empty." 假报错失败。复制成本一次性（MB 级 <2ms），远小于解析
// 成本；File/Blob 每次 arrayBuffer() 都是新读，无需复制。
async function toBytes(source: PdfSource): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source.slice();
  if (source instanceof ArrayBuffer) return new Uint8Array(source.slice());
  const buf = await source.arrayBuffer();
  return new Uint8Array(buf);
}

/** 用户回调异常隔离：调用方回调抛错不得被误报成渲染错误（word-preview 同口径）。 */
function guardCallback<A extends unknown[]>(
  cb: ((...args: A) => void) | undefined,
  name: string,
): ((...args: A) => void) | undefined {
  if (!cb) return cb;
  return (...args: A) => {
    try {
      cb(...args);
    } catch (e) {
      console.warn(`[pdf-preview] user ${name} callback threw; ignoring`, e);
    }
  };
}

/** 上游异常按 name 识别（vi.mock 场景 instanceof 失效，name 兜底）。 */
function errorName(err: unknown): string {
  return err instanceof Error ? err.name : "";
}

function isCancelError(err: unknown): boolean {
  const name = errorName(err);
  return (
    err instanceof RenderingCancelledException ||
    name === "RenderingCancelledException" ||
    name === "AbortException"
  );
}

/** worker/资产加载失败（而不是 PDF 内容损坏）的识别依据。 */
const WORKER_SETUP_ERROR_RE =
  /Setting up fake worker failed|No "GlobalWorkerOptions\.workerSrc" specified|Failed to fetch dynamically imported module|Invalid `workerSrc` type/i;

/** 解析/渲染失败归一为公开错误码（sniff 之后的失败路径共用）。 */
function normalizeError(err: unknown): PdfPreviewError {
  if (
    err instanceof PasswordException ||
    errorName(err) === "PasswordException"
  ) {
    // code 1 = NEED_PASSWORD，2 = INCORRECT_PASSWORD（上游 PasswordResponses）
    const code = (err as { code?: number }).code;
    return {
      code: "PASSWORD_PROTECTED",
      message:
        code === 2
          ? "The PDF is password-protected and the password was rejected. Encrypted PDFs are not supported."
          : "The PDF is password-protected. Encrypted PDFs are not supported.",
      cause: err,
    };
  }
  if (
    err instanceof InvalidPDFException ||
    errorName(err) === "InvalidPDFException"
  ) {
    return {
      code: "CORRUPT",
      message: "The PDF file is invalid or corrupted.",
      cause: err,
    };
  }
  const message = err instanceof Error ? err.message : String(err);
  // PDF 本身可能完全合法，只是 worker 资产不可达——归 UNKNOWN 并给出
  // 可操作的提示，而不是让用户误以为文件损坏（方案风险 7 的口径）。
  if (WORKER_SETUP_ERROR_RE.test(message)) {
    return {
      code: "UNKNOWN",
      message:
        `Failed to load the PDF worker: ${message} ` +
        "Check that pdf.worker.min.mjs is copied and reachable (workerSrc / asset copy).",
      cause: err,
    };
  }
  // MissingPDFException 是 url 源专用异常，6.x 已不再从主入口导出——
  // 内存字节源不可达，无需归一
  return {
    code: "CORRUPT",
    message: `Failed to parse or render the PDF: ${message}`,
    cause: err,
  };
}

export function createPdfPreview(
  container: HTMLElement,
  options: PdfPreviewOptions,
): PdfPreviewInstance {
  const t0 = performance.now();
  const onRendered = guardCallback(options.onRendered, "onRendered");
  const onError = guardCallback(options.onError, "onError");
  let destroyed = false;

  // 数值 scale 是编程期常量：非法值同步抛 TypeError（fail-fast），
  // 不进异步错误回调——fit-* 依赖容器量测，留给 boot 内解析
  if (typeof options.scale === "number") {
    if (!Number.isFinite(options.scale) || options.scale <= 0) {
      throw new TypeError(
        `Invalid scale: ${options.scale}. Use a positive number.`,
      );
    }
  } else if (
    options.scale !== undefined &&
    options.scale !== "fit-width" &&
    options.scale !== "fit-page"
  ) {
    // JS 调用方可能传任意字符串；不静默落入 fit-page 分支
    throw new TypeError(
      `Invalid scale: ${String(options.scale)}. Use a positive number, "fit-width" or "fit-page".`,
    );
  }

  ensurePdfStyles(options.styleNonce);

  // 私有 wrapper：destroy 同步摘除它即清空用户容器；迟到的渲染结果只会
  // 写入已脱离文档的节点（word-preview 同款防竞态结构）。
  const wrapper = document.createElement("div");
  wrapper.className = ROOT_CLASS;
  container.appendChild(wrapper);

  const pages: PageEntry[] = [];
  const queue: PageEntry[] = [];
  let loadingTask: PDFDocumentLoadingTask | null = null;
  let pdf: PDFDocumentProxy | null = null;
  let resolvedScale = 1;
  let firstReported = false;
  let rendering = false;
  /** 文档级失败终态：teardown 后所有迟到渲染/错误一律静默（destroy 另由
   * destroyed 表达；失败后仍允许消费方调用 destroy() 幂等收尾）。 */
  let failed = false;

  // 双观察器：renderIO 触发渲染，releaseIO 触发释放（两个 margin 的差值
  // 构成 1-3 屏的渲染保留缓冲）。root 取真正会滚动的祖先（无则回退视口），
  // rootMargin 才能按滚动容器高度展开预载带（见 findScrollRoot 注释）。
  const scrollRoot = findScrollRoot(container);
  const renderIO = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        const e = entry.target as HTMLDivElement & { __mop?: PageEntry };
        const page = e.__mop;
        if (!page) continue;
        // wanted 必须无条件更新（含 rendering 状态）：渲染被 release 取消后
        // 用户可能在取消落定前又滚回渲染带——此时忽略进入事件会让取消完成
        // 的页面停在 empty 且无人重新入队（滚动位置不再变化，IO 不会再回调）。
        page.wanted = entry.isIntersecting;
        if (entry.isIntersecting && page.state === "empty") {
          page.state = "queued";
          queue.push(page);
          pump();
        }
      }
    },
    { root: scrollRoot, rootMargin: RENDER_MARGIN },
  );

  const releaseIO = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        const e = entry.target as HTMLDivElement & { __mop?: PageEntry };
        const page = e.__mop;
        if (!page) continue;
        if (entry.isIntersecting) {
          page.pendingRelease = false;
        } else if (page.state === "rendered" || page.state === "rendering") {
          releaseEntry(page);
        }
      }
    },
    { root: scrollRoot, rootMargin: RELEASE_MARGIN },
  );

  /** 串行队列：同一时刻最多 1 页在渲（scroll 洪峰限流）；离带/过期条目直接丢弃。 */
  function pump(): void {
    if (rendering) return;
    for (;;) {
      const next = queue.shift();
      if (!next) return;
      if (next.state !== "queued") continue; // 过期条目（已释放/已渲）
      if (!next.wanted) {
        next.state = "empty"; // 入队后滚出渲染带：弃渲
        continue;
      }
      rendering = true;
      void renderEntry(next).finally(() => {
        rendering = false;
        pump();
      });
      return;
    }
  }

  function removeEntryDom(e: PageEntry): void {
    if (e.canvas) {
      // 置 0 立即释放 canvas bitmap 显存，再移除节点
      e.canvas.width = 0;
      e.canvas.height = 0;
      e.canvas.remove();
      e.canvas = null;
    }
    e.textDiv?.remove();
    e.textDiv = null;
  }

  /** 取消在途文本层：release/destroy 时文本流仍可能占用 worker 与主线程。 */
  function cancelEntryTextLayer(e: PageEntry): void {
    e.textLayer?.cancel();
    e.textLayer = null;
  }

  /** 释放主线程页资源（operator list / objs）；重入时上游按需重新获取。 */
  function cleanupPage(e: PageEntry): void {
    try {
      e.page?.cleanup();
    } catch (err) {
      // 清理失败不影响后续重新渲染，但不得静默吞掉
      console.warn("[pdf-preview] page cleanup failed; ignoring", err);
    }
  }

  /** rendered/rendering → empty：释放 canvas、文本层与页资源，重入按同路径重渲。 */
  function releaseEntry(e: PageEntry): void {
    if (e.state === "rendering") {
      e.pendingRelease = true; // 渲染完成后在 renderEntry 内收尾释放
      e.task?.cancel();
      cancelEntryTextLayer(e);
      // cleanup 在 renderTask 结束后由上游自动收尾；此处调用即登记
      cleanupPage(e);
      return;
    }
    if (e.state !== "rendered") return;
    cancelEntryTextLayer(e);
    removeEntryDom(e);
    e.state = "empty";
    e.pendingRelease = false;
    cleanupPage(e);
  }

  async function renderEntry(e: PageEntry): Promise<void> {
    e.state = "rendering";
    const tStart = performance.now();
    try {
      const page = e.page ?? (e.page = await pdf!.getPage(e.num));
      if (destroyed || failed) return;
      const viewport = page.getViewport({ scale: resolvedScale });
      const cssW = Math.max(1, Math.floor(viewport.width));
      const cssH = Math.max(1, Math.floor(viewport.height));
      e.el.style.width = `${cssW}px`;
      e.el.style.height = `${cssH}px`;
      // 文本层 CSS 变量链的页级输入（text-layer-css.ts 的 .mop-pdf-page 段）
      e.el.style.setProperty("--scale-factor", String(viewport.scale));

      const canvas = document.createElement("canvas");
      const dpr =
        typeof devicePixelRatio === "number" && devicePixelRatio > 0
          ? devicePixelRatio
          : 1;
      canvas.width = Math.max(1, Math.floor(viewport.width * dpr));
      canvas.height = Math.max(1, Math.floor(viewport.height * dpr));
      canvas.style.width = `${cssW}px`;
      canvas.style.height = `${cssH}px`;
      e.canvas = canvas;
      e.el.appendChild(canvas);

      const task = page.render({
        canvas,
        viewport,
        // 高分屏按 dpr 放大位图、CSS 尺寸不变（官方 viewer 同款做法）
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      });
      e.task = task;
      await task.promise;
      e.task = null;
      if (destroyed || failed) return;

      if (options.textLayer !== false) {
        const div = document.createElement("div");
        div.className = "textLayer";
        e.textDiv = div;
        e.el.appendChild(div);
        // 流式文本源：大文档首页出图不必等全文提取完成
        const tl = new TextLayer({
          textContentSource: page.streamTextContent(),
          container: div,
          viewport,
        });
        e.textLayer = tl;
        await tl.render();
        e.textLayer = null;
        if (destroyed || failed) return;
      }

      e.state = "rendered";
      if (!firstReported) {
        firstReported = true;
        const now = performance.now();
        onRendered?.({
          pageCount: pdf!.numPages,
          firstPageMs: Math.round(now - tStart),
          durationMs: Math.round(now - t0),
        });
      }
      // 释放请求到达后用户又滚回了渲染带（wanted=true）：保留本次渲染结果，
      // 不能按过期标记立即释放——否则刚回到视口的页面会瞬间变空白。
      if (e.pendingRelease && !e.wanted) releaseEntry(e);
      else e.pendingRelease = false;
    } catch (err) {
      e.task = null;
      cancelEntryTextLayer(e);
      if (destroyed || failed) return;
      if (isCancelError(err)) {
        // release/destroy 触发的取消：清场回 empty（重入视口可重渲），静默
        removeEntryDom(e);
        e.state = "empty";
        e.pendingRelease = false;
        cleanupPage(e);
        // 取消落定前页面已重新进入渲染带：立即重新入队（此时 pump 可能仍被
        // 本次 renderEntry 占据，finally 收尾后会再次 pump 取到该条目）。
        if (e.wanted) {
          e.state = "queued";
          queue.push(e);
          pump();
        }
        return;
      }
      removeEntryDom(e);
      e.state = "empty";
      // 失败分支同样要清 pendingRelease：残留标记会让下一次成功渲染在完成
      // 瞬间被误判为过期页立即释放。
      e.pendingRelease = false;
      cleanupPage(e);
      onError?.(normalizeError(err));
    }
  }

  function createPageEntry(num: number, page: PDFPageProxy): PageEntry {
    const el = document.createElement("div");
    el.className = PAGE_CLASS;
    const viewport = page.getViewport({ scale: resolvedScale });
    el.style.width = `${Math.max(1, Math.floor(viewport.width))}px`;
    el.style.height = `${Math.max(1, Math.floor(viewport.height))}px`;
    // 占位期即写变量：文本层样式链在渲染前就有正确输入（renderEntry
    // 渲染时会按最终 viewport 再写一次，同值幂等）
    el.style.setProperty("--scale-factor", String(resolvedScale));
    const entry: PageEntry = {
      num,
      el,
      page,
      state: "empty",
      wanted: false,
      pendingRelease: false,
      canvas: null,
      textDiv: null,
      textLayer: null,
      task: null,
    };
    // IO 回调按元素反查页条目（WeakMap 亦可；单一下挂属性更直接且不外泄——
    // 属性挂在包私有 DOM 上，序列化/消费方均不可见）
    (el as HTMLDivElement & { __mop?: PageEntry }).__mop = entry;
    pages.push(entry);
    wrapper.appendChild(el);
    renderIO.observe(el);
    releaseIO.observe(el);
    return entry;
  }

  /** 资源全量回收，destroy() 与文档级失败终态共用：断开观察器、取消在途
   * 渲染/文本层、清页引用与 DOM、终止 worker。失败终态调用它但不置
   * destroyed——消费方随后的 destroy() 仍幂等安全（worker 已在此终止，
   * 不会二次 destroy），迟到的渲染结果只会写入已脱离文档的节点。
   *
   * 注意 wrapper 被 destroy 闭包持有：只 remove() 会把已渲染的 canvas
   * 节点（及其位图）一起留在闭包里——必须先把子节点/位图显式释放。 */
  function teardown(): void {
    failed = true;
    renderIO.disconnect();
    releaseIO.disconnect();
    for (const e of pages) {
      e.task?.cancel();
      cancelEntryTextLayer(e);
      removeEntryDom(e);
      cleanupPage(e);
    }
    queue.length = 0;
    pages.length = 0;
    pdf = null;
    // 上游真实取消：终止 worker 与传输；destroy/失败路径的拒绝是预期
    // 噪音，静默消化（取消不是错误）
    void loadingTask?.destroy().catch(() => {});
    loadingTask = null;
    wrapper.replaceChildren();
    wrapper.remove();
  }

  const boot = async (): Promise<void> => {
    let bytes: Uint8Array;
    try {
      bytes = await toBytes(options.source);
    } catch (err) {
      if (destroyed) return;
      teardown();
      onError?.({
        code: "UNKNOWN",
        message: `Failed to read the source: ${err instanceof Error ? err.message : String(err)}`,
        cause: err,
      });
      return;
    }
    try {
      if (destroyed) return;

      const sniffed = sniffPdf(bytes);
      if (!sniffed.ok) {
        teardown();
        onError?.(sniffed.error);
        return;
      }
      if (destroyed) return;

      // 资产 URL 配置错误属调用方输入问题（UNKNOWN），不要落到 CORRUPT
      let urls: PdfAssetUrls;
      try {
        urls = resolveAssetUrls(options);
      } catch (err) {
        teardown();
        onError?.({
          code: "UNKNOWN",
          message: `Invalid asset URL configuration: ${err instanceof Error ? err.message : String(err)}`,
          cause: err,
        });
        return;
      }

      // Node（SSR import 后误用/测试环境）无 Worker 全局：不设 workerSrc，
      // 让上游走 fake worker 默认路径；浏览器必配 worker（上游缺配直接
      // 抛错，无静默 fallback——见包 README）。GlobalWorkerOptions 是模块
      // 级全局：多实例传不同 workerSrc 时后设者生效，属已知取舍。
      if (typeof Worker !== "undefined") {
        GlobalWorkerOptions.workerSrc = urls.workerSrc;
      }

      const ro = options.renderOptions ?? {};
      loadingTask = getDocument({
        data: bytes,
        cMapUrl: urls.cMapUrl,
        cMapPacked: true, // 随包分发的是 .bcmap 二进制格式
        standardFontDataUrl: urls.standardFontDataUrl,
        wasmUrl: options.wasmUrl,
        iccUrl: options.iccUrl,
        enableXfa: false, // XFA 表单首版不支持
        disableFontFace: ro.disableFontFace,
        useSystemFonts: ro.useSystemFonts,
      });
      pdf = await loadingTask.promise;
      if (destroyed) return;

      const first = await pdf.getPage(1);
      if (destroyed) return;

      // scale 解析：fit-* 依赖首页尺寸 + 容器量测；容器不可见（0 宽）退 1
      // （数值 scale 已在入口同步校验）
      const scaleOpt = options.scale ?? "fit-width";
      if (typeof scaleOpt === "number") {
        resolvedScale = scaleOpt;
      } else {
        const base = first.getViewport({ scale: 1 });
        const cw = container.clientWidth;
        if (!cw) {
          console.warn(
            "[pdf-preview] container has no measurable width (hidden?); falling back to scale=1.",
          );
          resolvedScale = 1;
        } else if (scaleOpt === "fit-width") {
          resolvedScale = cw / base.width;
        } else {
          const ch = container.clientHeight;
          resolvedScale = Math.min(
            cw / base.width,
            ch > 0 ? ch / base.height : Infinity,
          );
        }
      }

      // 首页占位即建、立即进入懒渲染管线（IO 初始回调会触发首页渲染）
      createPageEntry(1, first);

      // 其余页并行取 viewport 后一次性建占位（顺序稳定）：滚动总高一次
      // 成型，逐页真实尺寸防混合尺寸文档的滚动条跳动。失败（worker 中途
      // 崩溃等）按 CORRUPT 上报。
      const total = pdf.numPages;
      if (total > 1) {
        const nums = Array.from({ length: total - 1 }, (_, i) => i + 2);
        const proxies = await Promise.all(
          nums.map(async (n) => pdf!.getPage(n)),
        );
        if (destroyed) return;
        for (const p of proxies) createPageEntry(p.pageNumber, p);
      }
    } catch (err) {
      if (destroyed) return;
      teardown();
      onError?.(normalizeError(err));
    }
  };
  void boot();

  return {
    destroy() {
      if (destroyed) return;
      destroyed = true;
      teardown();
    },
  };
}
