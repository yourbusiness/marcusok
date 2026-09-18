import type { ExportOptions, ExportPhase, ExportResult } from "./types";
import { exportExcel } from "./index";
import { OverlayState, type OverlaySnapshot } from "./overlay-state";

/**
 * 导出期间的全局遮罩 + 进度条（可选能力，独立子路径
 * `@marcusok/excel-exporter/overlay`）。
 *
 * 主入口 (`./index`) 不引用本模块，因此不用这个能力的调用方打包体积不变。
 *
 * 进度来源是库已有的 `onProgress` / `onPhase`，本模块只做呈现。两条数据源
 * 的真实粒度必须清楚，否则会误以为是渲染 bug：
 * - 中间进度（0..1 之间）**只有 stream 路径**会产生，每 1000 行一次；
 * - main / worker 的 workbook 路由只有入口的 0 与收尾的 1。
 * 所以 workbook 路由会全程停留在不确定态（扫光动画），这是数据源粒度决定
 * 的，不是遮罩没生效。
 *
 * 主线程阻塞的限制：`fast-xlsx` 的构建与 `WorkbookBuilder.addSheet` 都是同步
 * 的，期间浏览器无法重绘。`delayMs > 0` 时若延迟定时器还没触发就被这次阻塞
 * 压住，遮罩将**完全不出现**（迟到的定时器会在 close 时被清掉，不会在导出
 * 结束后弹出来）。想在阻塞路由上也看到遮罩，用 `delayMs: 0`：此时遮罩在
 * `exportExcel` 之前同步挂载，紧接着的 nextPaint() 让它先绘制出来，代价是
 * 快导出会闪一下。无论哪种设置，阻塞期间动画由合成器继续推进，但百分比与
 * 文案会冻结到导出结束。
 */

/** 文案覆盖项；未提供的项用默认中文文案。 */
export interface OverlayTextOptions {
  title?: string;
  preparing?: string;
  building?: string;
  downloading?: string;
  finishing?: string;
  /** 不确定态下显示的补充说明（确定态自动隐藏）。 */
  hint?: string;
}

export interface OverlayOptions {
  /** 显示前延迟（ms），默认 200：导出在此之前结束则遮罩完全不出现。 */
  delayMs?: number;
  /**
   * 最短可见时长（ms），默认 300：已显示的遮罩关闭时若不足此值则延后移除，
   * 避免一闪而过。仅在遮罩真的显示过之后生效。
   */
  minVisibleMs?: number;
  /** 淡出时长（ms），默认 150。 */
  fadeOutMs?: number;
  /** 遮罩层级，默认 2147483000。 */
  zIndex?: number;
  /** 挂载容器，默认 document.body。 */
  container?: HTMLElement;
  /** 是否阻断页面交互，默认 true。设为 false 时只做视觉覆盖。 */
  blockInteraction?: boolean;
  /** 主题，默认 "auto"（按 prefers-color-scheme 解析）。 */
  theme?: "auto" | "light" | "dark";
  /** 文案覆盖。 */
  text?: OverlayTextOptions;
  /**
   * @internal 本路由是否会有 download 阶段，决定 build 之后的文案。由
   * `exportExcelWithOverlay` 自动推断；直接用 `showExportOverlay` 时默认 true。
   */
  willDownload?: boolean;
}

/** 遮罩句柄。所有方法都保证不抛错：遮罩是增强，不允许因它弄挂导出。 */
export interface ExportOverlayHandle {
  /** 接到 `ExportOptions.onProgress`。 */
  handleProgress(progress: number): void;
  /** 接到 `ExportOptions.onPhase`（第二参数未使用，可直接透传）。 */
  handlePhase(phase: ExportPhase): void;
  /** 关闭遮罩。幂等；挂在 finally 里即可，绝不能由 `onProgress === 1` 触发。 */
  close(): void;
}

interface ResolvedOptions {
  delayMs: number;
  minVisibleMs: number;
  fadeOutMs: number;
  zIndex: number;
  container: HTMLElement;
  blockInteraction: boolean;
  theme: "light" | "dark";
  willDownload: boolean;
  text: Required<OverlayTextOptions>;
}

const DEFAULT_TEXT: Required<OverlayTextOptions> = {
  title: "正在导出 Excel",
  preparing: "准备中…",
  building: "正在构建工作簿…",
  downloading: "正在下载…",
  finishing: "即将完成…",
  hint: "数据量较大时可能需要数十秒，请勿关闭页面",
};

const STYLE_ID = "mxe-overlay-style";

/**
 * 样式走一次注入的 <style> 而非全内联：`@keyframes` 与
 * `@media (prefers-reduced-motion)` 只能存在于样式表里，内联样式表达不了。
 * 动态部分（层级、透明度、确定态宽度）才用内联样式。
 *
 * 颜色用 CSS 变量承载，主题切换只改变量（auto 在挂载时由 JS 解析成
 * light/dark，因此这里只需一份 dark 覆盖，不必在媒体查询里重复一遍）。
 * 结构属性加 !important 是为了不被打包进宿主页面后撞上对方的全局样式。
 */
const STYLES = `
.mxe-overlay,
.mxe-overlay *,
.mxe-overlay *::before,
.mxe-overlay *::after {
  box-sizing: border-box;
}
@keyframes mxe-sweep {
  from { transform: translateX(-100%); }
  to   { transform: translateX(350%); }
}
.mxe-overlay {
  position: fixed !important;
  inset: 0 !important;
  display: flex !important;
  align-items: center;
  justify-content: center;
  opacity: 1;
  transition: opacity .15s ease-out;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC",
    "Microsoft YaHei", sans-serif;
  font-size: 14px;
  line-height: 1.5;
  text-align: left;
  /* 遮罩自持滚动拦截，不改 document.body.overflow：改 body 状态会在 close
     漏调时让宿主页面永久不可滚动，风险不对等。 */
  touch-action: none;
  overscroll-behavior: contain;
  --mxe-backdrop: rgba(255, 255, 255, .72);
  --mxe-panel-bg: #ffffff;
  --mxe-panel-border: rgba(0, 0, 0, .08);
  --mxe-panel-shadow: 0 6px 30px rgba(0, 0, 0, .12);
  --mxe-title: #1f2329;
  --mxe-text: #646a73;
  --mxe-track: #e5e6eb;
  --mxe-fill: #1677ff;
}
.mxe-overlay[data-mxe-theme="dark"] {
  --mxe-backdrop: rgba(0, 0, 0, .62);
  --mxe-panel-bg: #1f1f1f;
  --mxe-panel-border: rgba(255, 255, 255, .1);
  --mxe-panel-shadow: 0 6px 30px rgba(0, 0, 0, .5);
  --mxe-title: #e8e8e8;
  --mxe-text: #a3a3a3;
  --mxe-track: #3a3a3a;
  --mxe-fill: #3c89ff;
}
.mxe-overlay {
  background: var(--mxe-backdrop);
  -webkit-backdrop-filter: blur(1px);
  backdrop-filter: blur(1px);
}
.mxe-panel {
  min-width: 280px;
  max-width: 80vw;
  padding: 20px 24px;
  border-radius: 10px;
  background: var(--mxe-panel-bg);
  border: 1px solid var(--mxe-panel-border);
  box-shadow: var(--mxe-panel-shadow);
  text-align: center;
}
.mxe-title {
  color: var(--mxe-title);
  font-size: 15px;
  font-weight: 600;
}
.mxe-bar {
  position: relative;
  overflow: hidden;
  height: 6px;
  margin: 14px 0 10px;
  border-radius: 3px;
  background: var(--mxe-track);
}
.mxe-fill {
  position: absolute;
  inset: 0;
  transform: scaleX(0);
  transform-origin: left center;
  background: var(--mxe-fill);
  transition: transform .2s ease-out;
}
.mxe-sweep {
  position: absolute;
  top: 0;
  bottom: 0;
  left: 0;
  width: 40%;
  background: linear-gradient(90deg, transparent, var(--mxe-fill), transparent);
  animation: mxe-sweep 1.2s ease-in-out infinite;
}
.mxe-label {
  color: var(--mxe-text);
  font-size: 13px;
}
.mxe-hint {
  margin-top: 6px;
  color: var(--mxe-text);
  font-size: 12px;
}
.mxe-overlay[data-mxe-mode="determinate"] .mxe-sweep,
.mxe-overlay[data-mxe-mode="determinate"] .mxe-hint {
  display: none;
}
.mxe-overlay[data-mxe-mode="indeterminate"] .mxe-fill {
  display: none;
}
@media (prefers-reduced-motion: reduce) {
  .mxe-overlay,
  .mxe-fill {
    transition: none;
  }
  .mxe-sweep {
    animation: none;
    transform: none;
    opacity: .55;
  }
}
`;

function ensureStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = STYLES;
  document.head.appendChild(style);
}

interface OverlayDom {
  root: HTMLDivElement;
  bar: HTMLDivElement;
  fill: HTMLDivElement;
  label: HTMLDivElement;
}

function createDom(cfg: ResolvedOptions): OverlayDom {
  const root = document.createElement("div");
  root.className = "mxe-overlay";
  root.dataset.mxeTheme = cfg.theme;
  root.dataset.mxeMode = "indeterminate";
  root.setAttribute("aria-busy", "true");
  root.style.zIndex = String(cfg.zIndex);
  root.style.opacity = "0";
  if (!cfg.blockInteraction) root.style.pointerEvents = "none";
  // 滚动拦截挂在遮罩自身：节点随遮罩移除，不会在宿主页面上留下监听器。
  if (cfg.blockInteraction) {
    const block = (e: Event): void => e.preventDefault();
    root.addEventListener("wheel", block, { passive: false });
    root.addEventListener("touchmove", block, { passive: false });
  }

  const panel = document.createElement("div");
  panel.className = "mxe-panel";

  const title = document.createElement("div");
  title.className = "mxe-title";
  title.textContent = cfg.text.title;

  const bar = document.createElement("div");
  bar.className = "mxe-bar";
  // 不确定态不给 aria-valuenow（ARIA 规范：progressbar 缺 valuenow 即不确定）
  bar.setAttribute("role", "progressbar");
  bar.setAttribute("aria-valuemin", "0");
  bar.setAttribute("aria-valuemax", "100");

  const fill = document.createElement("div");
  fill.className = "mxe-fill";
  const sweep = document.createElement("div");
  sweep.className = "mxe-sweep";
  bar.append(fill, sweep);

  const label = document.createElement("div");
  label.className = "mxe-label";
  label.setAttribute("role", "status");
  label.setAttribute("aria-live", "polite");
  label.textContent = cfg.text.preparing;

  const hint = document.createElement("div");
  hint.className = "mxe-hint";
  hint.textContent = cfg.text.hint;

  panel.append(title, bar, label, hint);
  root.appendChild(panel);
  return { root, bar, fill, label };
}

function render(
  dom: OverlayDom,
  snap: OverlaySnapshot,
  text: Required<OverlayTextOptions>,
): void {
  const mode = snap.determinate === null ? "indeterminate" : "determinate";
  if (dom.root.dataset.mxeMode !== mode) dom.root.dataset.mxeMode = mode;
  if (snap.determinate !== null) {
    dom.fill.style.transform = `scaleX(${snap.determinate})`;
    dom.bar.setAttribute(
      "aria-valuenow",
      String(Math.round(snap.determinate * 100)),
    );
  } else {
    dom.bar.removeAttribute("aria-valuenow");
  }
  const next = text[snap.label];
  // 仅在变化时写：该节点是 aria-live 区域，重复写同一文本会被读屏重复播报。
  if (dom.label.textContent !== next) dom.label.textContent = next;
}

/**
 * 模块级单例：并发导出共用一份 DOM，`refs` 计数归零才移除。渲染以**最后更新
 * 者**为准（避免两个遮罩叠在一起相互遮挡）。
 */
interface ActiveOverlay {
  container: HTMLElement;
  dom: OverlayDom;
  refs: number;
  /** 当前驱动渲染的状态机（最后一次 show / progress / phase 的实例）。 */
  driver: OverlayState | null;
  /** 延迟显示定时器。close 时必须清掉，否则导出比 delayMs 快时它会迟到弹出。 */
  revealTimer: ReturnType<typeof setTimeout> | null;
  unmountTimer: ReturnType<typeof setTimeout> | null;
  fadeTimer: ReturnType<typeof setTimeout> | null;
  revealed: boolean;
}

let active: ActiveOverlay | null = null;

function clearTimer(timer: ReturnType<typeof setTimeout> | null): void {
  if (timer !== null) clearTimeout(timer);
}

function detach(h: ActiveOverlay): void {
  h.dom.root.remove();
  if (active === h) active = null;
}

function teardown(h: ActiveOverlay, fadeOutMs: number): void {
  // 淡出期间又来了新导出（refs > 0）→ 保留节点，由新的 show 重新接管。
  if (h.refs > 0) return;
  h.dom.root.style.opacity = "0";
  h.fadeTimer = setTimeout(() => {
    h.fadeTimer = null;
    if (h.refs > 0) return;
    detach(h);
  }, fadeOutMs);
}

function reveal(h: ActiveOverlay, text: Required<OverlayTextOptions>): void {
  h.revealTimer = null;
  if (h.refs <= 0 || h.revealed || !h.driver) return;
  h.revealed = true;
  render(h.dom, h.driver.reveal(), text);
  if (!h.dom.root.isConnected) h.container.appendChild(h.dom.root);
  // 先落 opacity:0 再强制样式计算，再置 1——不这样做首帧就是终值，过渡不生效。
  h.dom.root.style.opacity = "0";
  void h.dom.root.offsetHeight;
  h.dom.root.style.opacity = "1";
}

const NOOP_HANDLE: ExportOverlayHandle = {
  handleProgress: () => {},
  handlePhase: () => {},
  close: () => {},
};

function resolveTheme(theme: OverlayOptions["theme"]): "light" | "dark" {
  if (theme === "light" || theme === "dark") return theme;
  return typeof window !== "undefined" &&
    typeof window.matchMedia === "function" &&
    window.matchMedia("(prefers-color-scheme: dark)").matches
    ? "dark"
    : "light";
}

/**
 * 显示全局遮罩，返回驱动它的句柄。
 *
 * 与 {@link exportExcelWithOverlay} 的关系：后者是包了 `exportExcel` 的便捷
 * 封装，前者供 `exportTable` / `exportEcharts` / 自定义流程直接接线。
 *
 * Node / SSR（无 `document`）下返回全空实现的句柄，调用方无需分支。
 */
export function showExportOverlay(
  options: OverlayOptions = {},
): ExportOverlayHandle {
  let cfg: ResolvedOptions;
  try {
    if (typeof document === "undefined") return NOOP_HANDLE;
    cfg = {
      delayMs: Math.max(0, options.delayMs ?? 200),
      minVisibleMs: Math.max(0, options.minVisibleMs ?? 300),
      fadeOutMs: Math.max(0, options.fadeOutMs ?? 150),
      zIndex: options.zIndex ?? 2147483000,
      container: options.container ?? document.body,
      blockInteraction: options.blockInteraction ?? true,
      theme: resolveTheme(options.theme),
      willDownload: options.willDownload ?? true,
      text: { ...DEFAULT_TEXT, ...options.text },
    };
    if (!cfg.container) return NOOP_HANDLE;
  } catch {
    return NOOP_HANDLE;
  }

  const state = new OverlayState({
    delayMs: cfg.delayMs,
    minVisibleMs: cfg.minVisibleMs,
    willDownload: cfg.willDownload,
  });

  let closed = false;

  try {
    ensureStyles();

    // 容器变了而旧遮罩还挂着：直接丢弃（以最新容器为准）。
    if (active && active.container !== cfg.container) {
      clearTimer(active.revealTimer);
      clearTimer(active.unmountTimer);
      clearTimer(active.fadeTimer);
      detach(active);
    }

    if (!active) {
      active = {
        container: cfg.container,
        dom: createDom(cfg),
        refs: 0,
        driver: state,
        revealTimer: null,
        unmountTimer: null,
        fadeTimer: null,
        revealed: false,
      };
    }
    const h = active;
    // 上一轮挂起的移除（淡出 / 最短可见延时）取消：本次复用同一份 DOM。
    clearTimer(h.unmountTimer);
    h.unmountTimer = null;
    clearTimer(h.fadeTimer);
    h.fadeTimer = null;
    h.driver = state;
    h.refs += 1;

    if (h.refs === 1 && !h.revealed) {
      const remaining = state.delayRemaining();
      if (remaining <= 0) reveal(h, cfg.text);
      else
        h.revealTimer = setTimeout(() => {
          reveal(h, cfg.text);
        }, remaining);
    }

    return {
      handleProgress: (progress: number): void => {
        // 遮罩的任何异常都不允许影响导出：整块吞掉。
        try {
          if (closed || active !== h) return;
          h.driver = state;
          const snap = state.progress(progress);
          if (snap && h.revealed) render(h.dom, snap, cfg.text);
        } catch {
          /* 渲染失败不影响导出 */
        }
      },
      handlePhase: (phase: ExportPhase): void => {
        try {
          if (closed || active !== h) return;
          h.driver = state;
          const snap = state.phase(phase);
          if (snap && h.revealed) render(h.dom, snap, cfg.text);
        } catch {
          /* 渲染失败不影响导出 */
        }
      },
      close: (): void => {
        try {
          if (closed) return;
          closed = true;
          if (active !== h) return;
          h.refs -= 1;
          const wait = state.close();
          if (h.refs > 0) return;
          // 关键：从未显示过时必须清掉延迟定时器，否则它会在 close 之后触发
          // 并把遮罩弹出来（导出比 delayMs 还快时必然发生）。
          clearTimer(h.revealTimer);
          h.revealTimer = null;
          if (!h.revealed) {
            // 压根没插进 DOM，无需淡出
            h.driver = null;
            detach(h);
            return;
          }
          if (wait <= 0) teardown(h, cfg.fadeOutMs);
          else
            h.unmountTimer = setTimeout(() => {
              h.unmountTimer = null;
              teardown(h, cfg.fadeOutMs);
            }, wait);
        } catch {
          /* 关闭失败不影响导出 */
        }
      },
    };
  } catch {
    return NOOP_HANDLE;
  }
}

/** 让出一帧（实际两帧，确保样式与布局都已提交）。 */
function nextPaint(): Promise<void> {
  if (typeof requestAnimationFrame !== "function") return Promise.resolve();
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
  });
}

/**
 * `exportExcel` 的遮罩封装：导出期间显示全局遮罩，结束后（成功或失败）移除。
 *
 * 调用方原有的 `onProgress` / `onPhase` 被**链式追加**而非替换——它们通常还
 * 驱动着自有的指标面板。
 *
 * @example
 * ```ts
 * import { exportExcelWithOverlay } from '@marcusok/excel-exporter/overlay';
 *
 * const result = await exportExcelWithOverlay({
 *   filename: 'report',
 *   sheets: [...],
 * });
 * ```
 */
export async function exportExcelWithOverlay(
  options: ExportOptions,
  overlay: OverlayOptions = {},
): Promise<ExportResult> {
  // Node 无 document 时库不触发下载，build 之后的文案应落在 finishing。
  const willDownload =
    options.download !== false && typeof document !== "undefined";
  const handle = showExportOverlay({ ...overlay, willDownload });
  try {
    // 先让出一帧：遮罩挂载与紧接其后的同步构建若落在同一个任务里，浏览器永远
    // 不会绘制遮罩（用户只看到页面卡死）。delayMs 为 0 时这一步是遮罩能被看见
    // 的唯一保证。
    await nextPaint();
    return await exportExcel({
      ...options,
      onProgress: (progress) => {
        handle.handleProgress(progress);
        options.onProgress?.(progress);
      },
      onPhase: (phase, durationMs) => {
        handle.handlePhase(phase);
        options.onPhase?.(phase, durationMs);
      },
    });
  } finally {
    // 关闭时机只能在这里：收尾的 onProgress(1) 在失败路径同样会发（types.ts
    // 契约），成功后 download 还排在它之后，用进度值判定关闭都会出错。
    handle.close();
  }
}
