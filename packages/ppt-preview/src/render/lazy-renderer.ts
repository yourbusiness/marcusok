/**
 * 懒渲染状态机：双 IntersectionObserver + 串行队列 + 页导航。
 *
 * 结构上复用 pdf-preview 已验证的框架（渲染带 1 屏预载 / 释放带 3 屏、
 * 串行队列、离带条目出队即弃、canvas 置 0 释放位图），并按 PPT 引擎的
 * 能力差异调整（见 docs/ ppt-preview-plan.md 五节）：
 * - `renderSlide()` 没有单页 cancel 句柄：在途渲染只标记 wanted=false /
 *   pendingRelease，完成后清理，不能假装立即取消；
 * - 上游没有单页 cleanup：释放只清 canvas 位图与文本 DOM，引擎页模型
 *   由 presentation.destroy() 统一回收；
 * - 新增导航（goToSlide/next/prev）：当前页按滚动位置的最可见占位计算
 *   （按需测量，不依赖可能过期的 IO 缓存），滚动只操作本包识别的滚动
 *   容器，避免 scrollIntoView 连带滚动外层页面。
 */
import { PAGE_CLASS } from "./text-layer-css";

/** 渲染带（进入即入队渲染）：视口上下各预载 1 屏。 */
const RENDER_MARGIN = "100% 0px 100% 0px";
/** 保留带（离开才释放 canvas）：视口上下各 3 屏，与渲染带之间构成缓冲。 */
const RELEASE_MARGIN = "300% 0px 300% 0px";

export type SlideState = "empty" | "queued" | "rendering" | "rendered";

export interface SlideEntry {
  /** 0-based 页索引（原始页序，含隐藏页）。 */
  index: number;
  el: HTMLDivElement;
  state: SlideState;
  /** 当前在渲染带内（renderIO 判定）——出队时复核，离带即弃。 */
  wanted: boolean;
  /** 释放请求到达时仍在渲染：渲染完成后立即释放。 */
  pendingRelease: boolean;
  canvas: HTMLCanvasElement | null;
  textDiv: HTMLDivElement | null;
}

export interface LazyRendererCallbacks {
  /**
   * 渲染一页（canvas 已由本模块创建并挂载）。resolve = 渲染完成（含
   * 文本层），reject = 渲染失败（引擎原始异常，归一由 onSlideError 侧做）。
   */
  renderSlide(entry: SlideEntry, canvas: HTMLCanvasElement): Promise<void>;
  /** 页渲染失败上报（每页一次一报，多次重试多次报；error 为引擎原始异常）。 */
  onSlideError(entry: SlideEntry, error: unknown): void;
  /** 首张成功页（每实例最多一次）。 */
  onFirstRendered(firstSlideMs: number): void;
}

/**
 * 判断滚动祖先是否真的可能裁剪内容，而不是"随内容一起长高"的
 * overflow:auto 包装层。内容驱动包装层会让 IO root 退化为整份文档高度、
 * 懒渲染失效为全量渲染（pdf-preview 的 play 200 页 demo 实测教训）。
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

/** 探测容器最近的、真正会裁剪内容的滚动祖先（无则 null → 视口）。
 *
 * 调用时机在首个 createEntries **之后**（ensureObservers 惰性创建）：
 * "当前已溢出"这条启发式在容器为空时恒为 false，类定义/flex 解析高度
 * 的面板只有占位内容撑出溢出后才可被检出（pdf-preview 同款机制）。
 * 不能用 computed height 补检：真实浏览器对 auto 高度也返回 used px，
 * 会把内容驱动的包装层误判为定高。 */
export function findScrollRoot(start: HTMLElement): HTMLElement | null {
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

export class LazySlideRenderer {
  private readonly entries: SlideEntry[] = [];
  private readonly queue: SlideEntry[] = [];
  private renderIO: IntersectionObserver | null = null;
  private releaseIO: IntersectionObserver | null = null;
  private scrollRoot: HTMLElement | null = null;
  private rendering = false;
  private destroyed = false;
  private firstReported = false;

  constructor(
    private readonly container: HTMLElement,
    private readonly callbacks: LazyRendererCallbacks,
  ) {}

  /**
   * 双观察器惰性创建（首个 createEntries 之后）：构造时容器是空的，
   * findScrollRoot 的"已溢出"启发式必然落空——类定义固定高度 / flex
   * 解析高度的滚动面板（无内联样式、无 max-height）只有等占位内容撑出
   * 实际溢出才暴露。created 后不再重建（root 不随滚动漂移）；destroy
   * 幂等收口（null 安全）。IO 缺失的极旧环境/SSR 误用会在首次
   * createEntries 抛出，由 createPptPreview 的 boot 兜底归一上报。
   */
  private ensureObservers(): void {
    if (this.renderIO) return;
    this.scrollRoot = findScrollRoot(this.container);
    this.renderIO = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const e = (entry.target as HTMLDivElement & { __slide?: SlideEntry })
            .__slide;
          if (!e) continue;
          // wanted 必须无条件更新（含 rendering 状态）：渲染中被 releaseIO
          // 标记过期后，用户可能在渲染完成前又滚回渲染带——忽略进入事件会
          // 让完成收尾按过期的 pendingRelease 释放刚回到视口的页。
          e.wanted = entry.isIntersecting;
          if (entry.isIntersecting && e.state === "empty") {
            e.state = "queued";
            this.queue.push(e);
            this.pump();
          }
        }
      },
      { root: this.scrollRoot, rootMargin: RENDER_MARGIN },
    );
    this.releaseIO = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const e = (entry.target as HTMLDivElement & { __slide?: SlideEntry })
            .__slide;
          if (!e) continue;
          if (entry.isIntersecting) {
            e.pendingRelease = false;
          } else if (e.state === "rendered" || e.state === "rendering") {
            this.releaseEntry(e);
          }
        }
      },
      { root: this.scrollRoot, rootMargin: RELEASE_MARGIN },
    );
  }

  /** 为 0..count-1 建占位并挂接 IO（load 完成后一次性调用）。 */
  createEntries(count: number, cssWidth: number, cssHeight: number): void {
    for (let i = 0; i < count; i++) {
      const el = document.createElement("div");
      el.className = PAGE_CLASS;
      el.style.width = `${cssWidth}px`;
      el.style.height = `${cssHeight}px`;
      const entry: SlideEntry = {
        index: i,
        el,
        state: "empty",
        wanted: false,
        pendingRelease: false,
        canvas: null,
        textDiv: null,
      };
      (el as HTMLDivElement & { __slide?: SlideEntry }).__slide = entry;
      this.entries.push(entry);
      this.container.appendChild(el);
    }
    // 占位全部插入后再建观察器并 observe：内容已就位，"已溢出"启发式
    // 对类定义/flex 高度的滚动面板生效（见 ensureObservers 注释）
    this.ensureObservers();
    for (const e of this.entries) {
      this.renderIO!.observe(e.el);
      this.releaseIO!.observe(e.el);
    }
  }

  /** 串行队列：同一时刻最多 1 页在渲（scroll 洪峰限流）；离带/过期条目直接丢弃。 */
  private pump(): void {
    if (this.rendering) return;
    for (;;) {
      const next = this.queue.shift();
      if (!next) return;
      if (next.state !== "queued") continue; // 过期条目（已释放/已渲）
      if (!next.wanted) {
        next.state = "empty"; // 入队后滚出渲染带：弃渲
        continue;
      }
      this.rendering = true;
      void this.renderEntry(next).finally(() => {
        this.rendering = false;
        this.pump();
      });
      return;
    }
  }

  /** 释放 DOM 资源：canvas 置 0 释放位图后移除，文本层一并摘除。 */
  private removeEntryDom(e: SlideEntry): void {
    if (e.canvas) {
      e.canvas.width = 0;
      e.canvas.height = 0;
      e.canvas.remove();
      e.canvas = null;
    }
    e.textDiv?.remove();
    e.textDiv = null;
  }

  /** rendered/rendering → empty：释放 canvas 与文本 DOM，重入按同路径重渲。 */
  private releaseEntry(e: SlideEntry): void {
    if (e.state === "rendering") {
      // 无单页 cancel 句柄：标记过期，renderEntry 完成后收尾释放
      e.pendingRelease = true;
      e.wanted = false;
      return;
    }
    if (e.state !== "rendered") return;
    this.removeEntryDom(e);
    e.state = "empty";
    // 显式复位（与 renderEntry 失败分支对称）：否则标记残留到下一次渲染，
    // 成功后会被误释放——目前靠 releaseIO 的 isIntersecting 分支兜底清除，
    // 但那是隐式时序保证，不作为正确性依赖
    e.pendingRelease = false;
  }

  private async renderEntry(e: SlideEntry): Promise<void> {
    e.state = "rendering";
    const tStart = performance.now();
    // canvas 由本模块创建：尺寸/挂载统一，释放路径（置 0 + remove）可控
    const canvas = document.createElement("canvas");
    e.canvas = canvas;
    e.el.appendChild(canvas);
    try {
      await this.callbacks.renderSlide(e, canvas);
      if (this.destroyed) return;
      e.state = "rendered";
      if (!this.firstReported) {
        this.firstReported = true;
        this.callbacks.onFirstRendered(Math.round(performance.now() - tStart));
      }
      // 过期标记只在页面确实不在渲染带时生效；完成后已重新滚回（wanted）
      // 的页面必须保留本次渲染结果。
      if (e.pendingRelease && !e.wanted) this.releaseEntry(e);
      else e.pendingRelease = false;
    } catch (err) {
      if (this.destroyed) return;
      this.removeEntryDom(e);
      e.state = "empty"; // 回 empty：滚出再滚入可重试（IO 只在交叉变化时回调）
      e.pendingRelease = false;
      this.callbacks.onSlideError(e, err);
    }
  }

  /**
   * 当前最可见页（0-based）：按需遍历占位测量视口交叉面积取最大。
   * 导航是低频交互，一次性测量换实时准确（缓存的 IO ratio 会过期）。
   * 无页或全不可见时返回 -1。
   */
  currentVisibleIndex(): number {
    if (this.entries.length === 0 || this.destroyed) return -1;
    const viewTop = this.scrollRoot
      ? this.scrollRoot.getBoundingClientRect()
      : null;
    const vy = viewTop ? viewTop.top : 0;
    const vh = viewTop ? viewTop.height : window.innerHeight;
    let best = -1;
    let bestArea = 0;
    for (const e of this.entries) {
      const r = e.el.getBoundingClientRect();
      const h = Math.min(r.bottom, vy + vh) - Math.max(r.top, vy);
      const w =
        Math.min(r.right, viewTop ? viewTop.right : Infinity) -
        Math.max(r.left, viewTop ? viewTop.left : 0);
      const area = h > 0 && w > 0 ? h * w : 0;
      if (area > bestArea) {
        bestArea = area;
        best = e.index;
      }
    }
    return best;
  }

  /**
   * 滚动到指定页（页顶对齐视口顶）。只调整滚动容器的 scrollTop /
   * 窗口 scrollBy，不使用 scrollIntoView（会连带滚动外层文档站页面）。
   */
  scrollToIndex(index: number): void {
    const e = this.entries[index];
    if (!e || this.destroyed) return;
    if (this.scrollRoot) {
      const rootRect = this.scrollRoot.getBoundingClientRect();
      const elRect = e.el.getBoundingClientRect();
      // 相对修正（scrollTop += delta）而非绝对计算，避免对 wrapper 偏移
      // 的假设；瞬时跳转（不传 behavior:'smooth'）——导航语义是翻页直达
      this.scrollRoot.scrollTop += elRect.top - rootRect.top;
    } else {
      window.scrollBy({ top: e.el.getBoundingClientRect().top });
    }
  }

  /** 全量清场：断观察器、清队列、释放全部页 DOM。幂等。 */
  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.renderIO?.disconnect();
    this.releaseIO?.disconnect();
    this.renderIO = null;
    this.releaseIO = null;
    for (const e of this.entries) this.removeEntryDom(e);
    this.queue.length = 0;
    this.entries.length = 0;
  }
}
