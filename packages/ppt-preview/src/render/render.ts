/**
 * createPptPreview 的实现：对 @silurus/ooxml PptxPresentation（main 模式）
 * 的薄封装 + 连续滚动懒渲染。
 *
 * 封装职责（偏离上游默认/补足缺口的部分，均见包 README / limits）：
 * 1. sniff 前置（../sniff.ts）——空输入与非 ZIP/OLE2 容器提前归一；OLE2
 *    放行给引擎区分 legacy/加密/未知 CFB；
 * 2. source 防御性复制（../source.ts）——引擎解析会 transfer 输入 buffer，
 *    调用方原始字节不得直接进引擎；
 * 3. 错误归一（../errors.ts）——OoxmlError 8 码 / 两族资源限制 / Worker
 *    与 WASM 失败映射为 5 个公开错误码；
 * 4. 懒渲染状态机（lazy-renderer.ts）——双 IO + 串行队列 + 页导航；
 * 5. 文本层单趟收集（text-layer.ts）——renderSlide 的 onTextRun 在正式
 *    渲染同一趟收集 runs，不二次渲染；
 * 6. destroy 生命周期——加载中 destroy 作废旧结果（load 无 AbortSignal，
 *    迟到 resolve 后立即销毁引擎实例），渲染中 destroy 标记过期 + 迟到
 *    清理；destroy 后所有回调静默。
 *
 * 渲染形态：headless PptxPresentation + 每页一个 canvas，静态预览（不进
 * presentSlide 播放路径）；renderSlide 显式传 skipMediaControls:true——
 * 媒体只保留静态 poster/图形，不出现貌似可操作的播放控件。
 */
import { PptxPresentation } from "@silurus/ooxml/pptx";
import type { LoadOptions, PptxTextRunInfo } from "@silurus/ooxml/pptx";
import { resolveWasmUrl } from "../assets";
import { normalizeError } from "../errors";
import { sniffPpt } from "../sniff";
import { toBytes, toLoadBuffer } from "../source";
import type {
  PptPreviewError,
  PptPreviewInstance,
  PptPreviewOptions,
} from "../types";
import { LazySlideRenderer, type SlideEntry } from "./lazy-renderer";
import { ROOT_CLASS, ensurePptStyles } from "./text-layer-css";
import { attachTextLayer } from "./text-layer";

/** 用户回调异常隔离：调用方回调抛错不得被误报成渲染错误（word/pdf 同
 * 口径）。异步回调的 rejected promise 同步 catch 兜不住——必须显式
 * .catch，否则就是 unhandled rejection（TS 允许 async 函数赋给返回
 * void 的回调类型）。 */
function guardCallback<A extends unknown[]>(
  cb: ((...args: A) => void) | undefined,
  name: string,
): ((...args: A) => void) | undefined {
  if (!cb) return cb;
  const warn = (e: unknown): void => {
    console.warn(`[ppt-preview] user ${name} callback threw; ignoring`, e);
  };
  return (...args: A) => {
    try {
      const result = cb(...args) as unknown;
      // 按 thenable 判断而非 `instanceof Promise`：iframe 等跨 realm 的
      // Promise 不满足后者，会让异步回调的 rejection 漏成 unhandled。
      if (
        result !== null &&
        (typeof result === "object" || typeof result === "function") &&
        typeof (result as { then?: unknown }).then === "function"
      ) {
        void (result as PromiseLike<unknown>).then(undefined, warn);
      }
    } catch (e) {
      warn(e);
    }
  };
}

export function createPptPreview(
  container: HTMLElement,
  options: PptPreviewOptions,
): PptPreviewInstance {
  const t0 = performance.now();
  const onRendered = guardCallback(options.onRendered, "onRendered");
  const onError = guardCallback(options.onError, "onError");
  let destroyed = false;
  /** 引擎实例；destroy 后由 engineDestroyed 标记表达销毁语义。 */
  let engine: PptxPresentation | null = null;
  let engineDestroyed = false;
  let renderer: LazySlideRenderer | null = null;
  let loadedCount = 0;
  let cssWidth = 0;
  let cssHeight = 0;
  /** 在途的引擎 load promise（destroy 的作废路径使用）。 */
  let pendingLoad: Promise<PptxPresentation> | null = null;

  // slideWidth / dpr 是编程期常量：非法值同步抛 TypeError（fail-fast），
  // 不进异步错误回调
  const slideWidthOpt = options.slideWidth ?? 960;
  if (
    typeof slideWidthOpt !== "number" ||
    !Number.isFinite(slideWidthOpt) ||
    slideWidthOpt <= 0
  ) {
    throw new TypeError(
      `Invalid slideWidth: ${String(slideWidthOpt)}. Use a positive finite number.`,
    );
  }
  const dprOpt = options.dpr;
  if (
    dprOpt !== undefined &&
    (typeof dprOpt !== "number" || !Number.isFinite(dprOpt) || dprOpt <= 0)
  ) {
    throw new TypeError(
      `Invalid dpr: ${String(dprOpt)}. Use a positive finite number.`,
    );
  }
  const dpr =
    dprOpt ??
    (typeof devicePixelRatio === "number" && devicePixelRatio > 0
      ? devicePixelRatio
      : 1);
  // 非法 cjkFallback 原本在 load() 内部才由上游抛 TypeError，最终被归一成
  // 异步 UNKNOWN（"Failed to load…"），与 slideWidth/dpr 的 fail-fast 口径
  // 不一致，也误导成文件/环境问题。合法集合与上游 LoadOptions 对齐。
  const cjkFallback = options.renderOptions?.cjkFallback;
  if (
    cjkFallback !== undefined &&
    !["auto", "kr", "sc", "tc", "hk", "jp"].includes(cjkFallback)
  ) {
    throw new TypeError(
      `Invalid cjkFallback: ${String(cjkFallback)}. Use "auto", "kr", "sc", "tc", "hk" or "jp".`,
    );
  }

  ensurePptStyles(options.styleNonce);

  // 私有 wrapper：destroy 同步摘除它即清空用户容器；迟到的渲染结果只会
  // 写入已脱离文档的节点（word/pdf-preview 同款防竞态结构）
  const wrapper = document.createElement("div");
  wrapper.className = ROOT_CLASS;
  container.appendChild(wrapper);

  /** 引擎销毁（幂等）：destroy() 与失败终态共用；异常静默（销毁路径不报错）。
   * 即使 engineDestroyed 已置位也要清空 engine 引用：加载中 destroy 时
   * destroy 的迟到 resolve 回调与 boot 的后续检查会各自拿到同一实例，先到
   * 者完成销毁，后者仍须释放引用，不能把已销毁的引擎留在闭包里。 */
  function killEngine(): void {
    const target = engine;
    engine = null;
    if (!target || engineDestroyed) return;
    engineDestroyed = true;
    try {
      target.destroy();
    } catch (err) {
      console.warn("[ppt-preview] engine destroy threw; ignoring", err);
    }
  }

  /** 失败终态：清场 + 作废页数 + 上报（destroy 触发的路径不进入此处）。
   * 消费方随后的 destroy() 仍幂等安全。 */
  function fail(error: PptPreviewError): void {
    loadedCount = 0;
    renderer?.destroy();
    renderer = null;
    wrapper.replaceChildren();
    wrapper.remove();
    killEngine();
    onError?.(error);
  }

  /**
   * 单页渲染（lazy-renderer 回调）：引擎 renderSlide + 单趟文本层。
   * 失败原样抛出（引擎异常），归一集中在 onSlideError 回调边界做。
   */
  async function renderOneSlide(
    entry: SlideEntry,
    canvas: HTMLCanvasElement,
  ): Promise<void> {
    // onTextRun 在正式渲染同一趟收集 runs（不调用 collectSlideRuns 的
    // 二次渲染路径）；文本层关闭时不收集
    const runs: PptxTextRunInfo[] = [];
    // engine! 的非空前提：本函数只能经 renderer 的串行队列到达，而 renderer
    // 创建于 engine 赋值之后；engine 置空的唯一路径 killEngine() 在
    // destroy()/fail() 中均于 renderer 已先行销毁后调用，队列随之清空，
    // 不会再有新的渲染发出（在途调用早已捕获引擎引用，不受置空影响）
    await engine!.renderSlide(canvas, entry.index, {
      width: cssWidth,
      dpr,
      skipMediaControls: true,
      ...(options.enableTextSelection === false
        ? {}
        : { onTextRun: (r: PptxTextRunInfo) => runs.push(r) }),
    });
    // 文本层挂在占位 el 上、覆盖 canvas（上游定位全内联，容器样式见
    // text-layer-css.ts）；textDiv 引用由 lazy-renderer 的释放路径清理
    if (options.enableTextSelection !== false) {
      entry.textDiv = attachTextLayer(entry.el, runs, cssWidth, cssHeight);
    }
  }

  const boot = async (): Promise<void> => {
    let bytes: Uint8Array;
    try {
      bytes = await toBytes(options.source);
    } catch (err) {
      if (destroyed) return;
      fail({
        code: "UNKNOWN",
        message: `Failed to read the source: ${err instanceof Error ? err.message : String(err)}`,
        cause: err,
      });
      return;
    }
    if (destroyed) return;

    const sniffed = sniffPpt(bytes);
    if (!sniffed.ok) {
      fail(sniffed.error);
      return;
    }

    // 资产 URL 配置错误属调用方输入问题（UNKNOWN），不要落到 CORRUPT
    let wasmUrl: string | URL | undefined;
    try {
      wasmUrl = resolveWasmUrl(options.wasmUrl);
    } catch (err) {
      fail({
        code: "UNKNOWN",
        message: `Invalid asset URL configuration: ${err instanceof Error ? err.message : String(err)}`,
        cause: err,
      });
      return;
    }

    const ro = options.renderOptions ?? {};
    const loadOpts: LoadOptions = {
      mode: "main", // 解析在 module Web Worker、渲染在主线程（本包固定形态）
      password: options.password,
      cjkFallback: ro.cjkFallback,
      resourceLimits: ro.resourceLimits,
      workerTimeoutMs: ro.workerTimeoutMs,
      useGoogleFonts: options.useGoogleFonts ?? false,
      wasmUrl,
    };

    let pres: PptxPresentation;
    try {
      // load() 正常返回 Promise，但非法选项等编程期错误可能同步抛出——
      // 必须在 try 内调用，否则 void boot() 会漏成 unhandled rejection。
      pendingLoad = PptxPresentation.load(toLoadBuffer(bytes), loadOpts);
      pres = await pendingLoad;
    } catch (err) {
      pendingLoad = null;
      if (destroyed) return;
      fail(normalizeError(err));
      return;
    }
    pendingLoad = null;
    // 后半段（引擎 getter 读取 / LazySlideRenderer 构造含双 IO / 建占位）
    // 属同步初始化，同样可能抛（IO 缺失的极旧环境、SSR 误用）：整体兜底
    // 经 fail 清场并归一上报，避免 void boot() 沦为 unhandled rejection、
    // 用户无任何反馈（pdf-preview 的 boot 为同款整体收尾）
    try {
      engine = pres;
      if (destroyed) {
        // 加载期间被 destroy：迟到结果立即销毁引擎，静默（取消不是错误）
        killEngine();
        return;
      }

      // 空 deck：上游接受（spike S3 实测）——slideCount=0、不触发 onRendered、
      // 不报错；wrapper 留空，destroy 摘除
      const count = pres.slideCount;
      if (count === 0) return;

      // deck 比例（EMU 原始值）：无效时按内容错误上报，不生成 NaN 占位
      const sw = pres.slideWidth;
      const sh = pres.slideHeight;
      const ratio = sh / sw;
      if (
        !Number.isFinite(sw) ||
        !Number.isFinite(sh) ||
        sw <= 0 ||
        sh <= 0 ||
        !Number.isFinite(ratio) ||
        ratio <= 0
      ) {
        fail({
          code: "CORRUPT",
          message: `The presentation reports an invalid slide size (${sw}x${sh}).`,
        });
        return;
      }
      cssWidth = Math.max(1, Math.floor(slideWidthOpt));
      cssHeight = Math.max(1, Math.round(cssWidth * ratio));
      loadedCount = count;

      renderer = new LazySlideRenderer(wrapper, {
        renderSlide: renderOneSlide,
        onSlideError: (_entry, error) => {
          // 归一集中点：renderOneSlide 抛的是引擎原始异常
          onError?.(normalizeError(error));
        },
        onFirstRendered: (firstSlideMs) => {
          const now = performance.now();
          onRendered?.({
            slideCount: loadedCount,
            firstSlideMs,
            durationMs: Math.round(now - t0),
          });
        },
      });
      // 建占位即进入懒渲染管线（IO 初始回调触发首页渲染）。占位包含隐藏页、
      // 保持原始页序（首版策略：全部显示，见 types.ts slideCount 注释）
      renderer.createEntries(count, cssWidth, cssHeight);
    } catch (err) {
      // destroy 已接管清理：迟到异常不再上报（取消不是错误）
      if (destroyed) return;
      fail(normalizeError(err));
    }
  };
  void boot();

  return {
    destroy() {
      if (destroyed) return;
      destroyed = true;
      // 渲染侧先停（在途渲染标记过期，迟到结果不写回调）；引擎随后销毁，
      // 在途 renderSlide 的 reject 由 lazy-renderer 的 destroyed 检查静默
      renderer?.destroy();
      renderer = null;
      wrapper.replaceChildren();
      wrapper.remove();
      if (engine) {
        killEngine();
      } else if (pendingLoad) {
        // 加载在途：无 AbortSignal，作废旧结果；迟到 resolve 立即销毁引擎，
        // 迟到 reject 静默（取消不是错误）
        const stale = pendingLoad;
        void stale
          .then((p) => {
            engine = p;
            killEngine();
          })
          .catch(() => {});
      }
    },
    goToSlide(index) {
      if (!Number.isInteger(index)) {
        throw new TypeError(
          `Invalid slide index: ${String(index)}. Use an integer (0-based).`,
        );
      }
      // 加载前/destroy 后：占位尚不存在，no-op（混合边界见 types.ts）
      if (destroyed || !renderer || loadedCount === 0) return;
      if (index < 0 || index >= loadedCount) {
        throw new RangeError(
          `Slide index ${index} is out of range (0..${loadedCount - 1}).`,
        );
      }
      renderer.scrollToIndex(index);
    },
    nextSlide() {
      if (destroyed || !renderer || loadedCount === 0) return;
      const cur = renderer.currentVisibleIndex();
      if (cur < 0 || cur + 1 >= loadedCount) return;
      renderer.scrollToIndex(cur + 1);
    },
    prevSlide() {
      if (destroyed || !renderer || loadedCount === 0) return;
      const cur = renderer.currentVisibleIndex();
      if (cur <= 0) return;
      renderer.scrollToIndex(cur - 1);
    },
    get slideCount() {
      if (destroyed || !renderer) return 0;
      return loadedCount;
    },
  };
}
