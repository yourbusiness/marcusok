import { OverlayState, type OverlaySnapshot } from "./overlay-state";

/**
 * 通用全屏进度遮罩：毛玻璃面板 + 旋转圆环（不确定态）/ 细进度条与百分比
 * （确定态），light/dark 主题，一次注入的样式，模块级单例共享。
 *
 * 本包是 @marcusok 各业务包（excel-exporter 及未来的导出/预览包）共用的
 * UI 层，不含任何业务语义：阶段 key 与文案由调用方定义，进度协议只有
 * `setProgress(0..1)` 与 `setPhase(key)` 两个入口。
 *
 * 进度粒度决定观感：只有会发出 (0,1) 中间值的调用方才会进入确定态，
 * 其余全程停留在不确定态（spinner）——这是数据源粒度，不是遮罩故障。
 *
 * 主线程阻塞的限制：调用方若在遮罩揭示前同步阻塞主线程（如同步构建
 * 大对象），`delayMs > 0` 的揭示定时器会被压住，遮罩**完全不出现**
 * （迟到的定时器会在 close 时被清掉，不会在任务结束后弹出来）。想在
 * 阻塞流程上也看到遮罩，用 `delayMs: 0`：遮罩在阻塞开始前同步挂载，
 * 紧接着的 nextPaint() 让它先绘制出来，代价是快任务会闪一下。无论哪种
 * 设置，阻塞期间动画由合成器继续推进，但百分比与文案会冻结到任务结束。
 */

/** 文案覆盖项；未提供的项用默认中文文案。 */
export interface ProgressOverlayTextOptions {
  /** 面板标题。 */
  title?: string;
  /** 收到首个阶段事件前的初始文案。 */
  initial?: string;
  /** 阶段 key → 文案映射；`setPhase(key)` 时查表，未覆盖的 key 原样显示。 */
  phases?: Record<string, string>;
  /** 不确定态下显示的补充说明（确定态自动隐藏）。 */
  hint?: string;
}

export interface ProgressOverlayOptions {
  /** 显示前延迟（ms），默认 200：任务在此之前结束则遮罩完全不出现。 */
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
  text?: ProgressOverlayTextOptions;
}

/** 遮罩句柄。所有方法都保证不抛错：遮罩是增强，不允许因它弄挂宿主任务。 */
export interface ProgressOverlayHandle {
  /** 进度（0..1）：中间值一到即切确定态。 */
  setProgress(progress: number): void;
  /** 阶段 key：查 `text.phases` 切换文案，未知 key 原样显示。 */
  setPhase(key: string): void;
  /** 关闭遮罩。幂等；挂在 finally 里即可，不得由 `setProgress(1)` 触发。 */
  close(): void;
}

interface ResolvedText {
  title: string;
  initial: string;
  phases: Record<string, string>;
  hint: string;
}

interface ResolvedOptions {
  delayMs: number;
  minVisibleMs: number;
  fadeOutMs: number;
  zIndex: number;
  container: HTMLElement;
  blockInteraction: boolean;
  theme: "light" | "dark";
  text: ResolvedText;
}

const DEFAULT_TEXT: ResolvedText = {
  title: "请稍候",
  initial: "正在处理…",
  phases: {},
  hint: "任务可能需要一些时间，请勿关闭页面",
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
@keyframes mxe-spin {
  to { transform: rotate(360deg); }
}
@keyframes mxe-pop {
  from {
    opacity: 0;
    transform: scale(.96) translateY(4px);
  }
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
  /* 中性灰阶（zinc 系）而非品牌蓝：遮罩不挑宿主站点配色，深浅两态只换变量 */
  --mxe-backdrop: rgba(82, 82, 91, .32);
  --mxe-panel-bg: rgba(255, 255, 255, .72);
  --mxe-panel-border: rgba(0, 0, 0, .06);
  --mxe-panel-shadow:
    inset 0 1px 0 rgba(255, 255, 255, .5),
    0 2px 8px rgba(0, 0, 0, .06),
    0 16px 48px rgba(0, 0, 0, .14);
  --mxe-title: #18181b;
  --mxe-text: #71717a;
  --mxe-hint: #a1a1aa;
  --mxe-track: rgba(0, 0, 0, .08);
  --mxe-fill: #171717;
}
.mxe-overlay[data-mxe-theme="dark"] {
  --mxe-backdrop: rgba(0, 0, 0, .55);
  --mxe-panel-bg: rgba(24, 24, 27, .65);
  --mxe-panel-border: rgba(255, 255, 255, .08);
  --mxe-panel-shadow:
    inset 0 1px 0 rgba(255, 255, 255, .06),
    0 2px 8px rgba(0, 0, 0, .4),
    0 16px 48px rgba(0, 0, 0, .55);
  --mxe-title: #fafafa;
  --mxe-text: #a1a1aa;
  --mxe-hint: #71717a;
  --mxe-track: rgba(255, 255, 255, .14);
  --mxe-fill: #fafafa;
}
.mxe-overlay {
  background: var(--mxe-backdrop);
  -webkit-backdrop-filter: blur(10px);
  backdrop-filter: blur(10px);
}
/* 玻璃面板：半透明底 + 大半径模糊 + 饱和提升；顶部内高光并入 shadow 变量。
   root 与 panel 各挂一层 backdrop-filter（前者虚化页面、后者造玻璃质感），
   中低端机有开销，但节点只在任务期间存在，量级可接受。入场微动效与 root
   的 opacity 淡入同帧起步（reveal 在挂载后下一帧才置 1）。 */
.mxe-panel {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 8px;
  min-width: 280px;
  max-width: 80vw;
  padding: 26px 30px;
  border-radius: 16px;
  background: var(--mxe-panel-bg);
  border: 1px solid var(--mxe-panel-border);
  box-shadow: var(--mxe-panel-shadow);
  -webkit-backdrop-filter: blur(24px) saturate(1.8);
  backdrop-filter: blur(24px) saturate(1.8);
  text-align: center;
  animation: mxe-pop .16s cubic-bezier(.16, 1, .3, 1) both;
}
.mxe-spinner {
  width: 28px;
  height: 28px;
  /* 旋转 svg 元素本身：transform-origin 默认即元素盒中心 */
  animation: mxe-spin .8s linear infinite;
}
.mxe-spinner-track {
  fill: none;
  stroke: var(--mxe-track);
  stroke-width: 2.5;
}
.mxe-spinner-arc {
  fill: none;
  stroke: var(--mxe-fill);
  stroke-width: 2.5;
  stroke-linecap: round;
}
.mxe-title {
  color: var(--mxe-title);
  font-size: 14px;
  font-weight: 500;
}
.mxe-bar {
  align-self: stretch;
  position: relative;
  overflow: hidden;
  height: 4px;
  border-radius: 999px;
  background: var(--mxe-track);
}
.mxe-fill {
  position: absolute;
  inset: 0;
  transform: scaleX(0);
  transform-origin: left center;
  background: var(--mxe-fill);
  transition: transform .2s cubic-bezier(.16, 1, .3, 1);
}
/* 文案行：label 居左、百分比居右；不确定态下 percent 被 display:none，
   行内只剩 label */
.mxe-row {
  align-self: stretch;
  display: flex;
  justify-content: space-between;
  gap: 12px;
}
.mxe-label {
  color: var(--mxe-text);
  font-size: 13px;
}
.mxe-percent {
  color: var(--mxe-text);
  font-size: 12px;
  /* 等宽数字：两位数与一位数切换时不抖动 */
  font-variant-numeric: tabular-nums;
}
.mxe-hint {
  color: var(--mxe-hint);
  font-size: 12px;
}
.mxe-overlay[data-mxe-mode="determinate"] .mxe-spinner,
.mxe-overlay[data-mxe-mode="determinate"] .mxe-hint {
  display: none;
}
.mxe-overlay[data-mxe-mode="indeterminate"] .mxe-bar,
.mxe-overlay[data-mxe-mode="indeterminate"] .mxe-percent {
  display: none;
}
@media (prefers-reduced-motion: reduce) {
  .mxe-overlay,
  .mxe-fill {
    transition: none;
  }
  .mxe-spinner,
  .mxe-panel {
    animation: none;
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
  title: HTMLDivElement;
  hint: HTMLDivElement;
  percent: HTMLDivElement;
}

const SVG_NS = "http://www.w3.org/2000/svg";

/**
 * 不确定态的旋转圆环：静态轨道圆 + 270° 圆头弧。用 SVG 而非 border 圆环，
 * 因为圆头端点与精确弧长（dasharray）在 border 技法上表达不了。
 */
function createSpinner(): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, "svg");
  svg.setAttribute("class", "mxe-spinner");
  svg.setAttribute("viewBox", "0 0 28 28");
  // 纯装饰：语义由 label 的 aria-live 状态区域承载
  svg.setAttribute("aria-hidden", "true");
  const circle = (cls: string, dash?: string): void => {
    const c = document.createElementNS(SVG_NS, "circle");
    c.setAttribute("class", cls);
    c.setAttribute("cx", "14");
    c.setAttribute("cy", "14");
    c.setAttribute("r", "11.5");
    if (dash !== undefined) c.setAttribute("stroke-dasharray", dash);
    svg.appendChild(c);
  };
  // 周长 2π·11.5 ≈ 72.3；弧长画 75%（270°）≈ 54.2，空段给足避免首尾相接
  circle("mxe-spinner-track");
  circle("mxe-spinner-arc", "54.2 72.3");
  return svg;
}

/**
 * 建窗/揭示路径上的渲染兜底：show 流程在 `refs += 1`、driver 入栈**之后**才
 * 同步调用 render/reveal，若让渲染异常逃出建窗 try，外层 catch 会返回
 * NOOP_HANDLE——调用方的 close() 变成空操作，refs 永远归零不了，已揭示的
 * 遮罩永久挂屏并持续拦截交互。渲染失败只影响当帧内容（后续 progress/phase
 * 事件会自愈重渲染），绝不许打断建窗。
 */
function renderSafely(run: () => void): void {
  try {
    run();
  } catch {
    /* 渲染失败不影响宿主任务，也不破坏建窗 */
  }
}

/**
 * 当前驱动渲染的状态机与其文案，**配对不可拆**：并发任务共用一份遮罩、
 * 以最后更新者为准，但揭开的定时器（revealTimer）可能属于另一个已 close
 * 的调用方——渲染时若取定时器发起者的文案，会把前一个任务的 title/hint
 * 泄漏给当前驱动者。状态与文案绑成一体后，无论哪个定时器触发 reveal，
 * 用的都是当前 driver 自己的文案。
 */
interface OverlayDriver {
  state: OverlayState;
  text: ResolvedText;
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
  bar.appendChild(fill);

  // label 与 percent 同行分居两端；percent 独立于 label 节点——label 是
  // aria-live 区域，混入百分比会让读屏把每次进度变化重播一遍。
  const row = document.createElement("div");
  row.className = "mxe-row";

  const label = document.createElement("div");
  label.className = "mxe-label";
  label.setAttribute("role", "status");
  label.setAttribute("aria-live", "polite");
  label.textContent = cfg.text.initial;

  const percent = document.createElement("div");
  percent.className = "mxe-percent";
  row.append(label, percent);

  const hint = document.createElement("div");
  hint.className = "mxe-hint";
  hint.textContent = cfg.text.hint;

  panel.append(createSpinner(), title, bar, row, hint);
  root.appendChild(panel);
  return { root, bar, fill, label, title, hint, percent };
}

/** 快照里的阶段 key → 实际显示文案；未覆盖的 key 原样显示（可发现拼写错误）。 */
function phaseText(text: ResolvedText, phase: string | null): string {
  if (phase === null) return text.initial;
  return text.phases[phase] ?? phase;
}

function render(
  dom: OverlayDom,
  snap: OverlaySnapshot,
  text: ResolvedText,
): void {
  const mode = snap.determinate === null ? "indeterminate" : "determinate";
  if (dom.root.dataset.mxeMode !== mode) dom.root.dataset.mxeMode = mode;
  if (snap.determinate !== null) {
    dom.fill.style.transform = `scaleX(${snap.determinate})`;
    dom.bar.setAttribute(
      "aria-valuenow",
      String(Math.round(snap.determinate * 100)),
    );
    const pct = `${Math.round(snap.determinate * 100)}%`;
    if (dom.percent.textContent !== pct) dom.percent.textContent = pct;
  } else {
    dom.bar.removeAttribute("aria-valuenow");
  }
  // title/hint 与 label 一样随当前 driver 刷新：DOM 是并发任务共享的单例，
  // createDom 时写入的是创建者的文案，接管者若不重写，后来者会一直看着
  // 前一个任务的标题。同样仅在变化时写，避免无谓的 DOM 变更。
  if (dom.title.textContent !== text.title) dom.title.textContent = text.title;
  if (dom.hint.textContent !== text.hint) dom.hint.textContent = text.hint;
  const next = phaseText(text, snap.phase);
  // 仅在变化时写：该节点是 aria-live 区域，重复写同一文本会被读屏重复播报。
  if (dom.label.textContent !== next) dom.label.textContent = next;
}

/**
 * 模块级单例：并发任务共用一份 DOM，`refs` 计数归零才移除。渲染以**最后更新
 * 者**为准（避免两个遮罩叠在一起相互遮挡）。
 */
interface ActiveOverlay {
  container: HTMLElement;
  dom: OverlayDom;
  refs: number;
  /** 当前驱动渲染的状态机+文案（最后一次 show / progress / phase 的实例）。 */
  driver: OverlayDriver | null;
  /**
   * 按介入顺序的调用方栈（show 压入、close 摘除自己）。并发 close 早退
   * （refs>0）时 driver 按栈回退到最后介入的仍活动调用方：挂起的揭示
   * 定时器可能正是早退者发起的，到点揭开时显示的必须是仍活动者的文案。
   */
  driverStack: OverlayDriver[];
  /**
   * 挂起的揭示定时器的截止时刻（Date.now 基准）。并发后来者的 delayMs 与
   * 已挂定时器的截止取更晚者——任一活动调用方的 delay 期未过都不揭示。
   */
  revealDeadline: number | null;
  /** 延迟显示定时器。close 时必须清掉，否则任务比 delayMs 快时它会迟到弹出。 */
  revealTimer: ReturnType<typeof setTimeout> | null;
  unmountTimer: ReturnType<typeof setTimeout> | null;
  fadeTimer: ReturnType<typeof setTimeout> | null;
  revealed: boolean;
  /**
   * 遮罩实际揭示的时刻（DOM 级）。minVisible 按它计算而非各 driver 自己的
   * state.revealedAt——并发下 reveal 记录在揭示当时的 driver 上，末位
   * close 者直取自己的 state 会得出 0 而提前拆台（一闪而过）。
   */
  revealedAt: number | null;
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
  // 淡出期间又来了新任务（refs > 0）→ 保留节点，由新的 show 重新接管。
  if (h.refs > 0) return;
  h.dom.root.style.opacity = "0";
  h.fadeTimer = setTimeout(() => {
    h.fadeTimer = null;
    if (h.refs > 0) return;
    detach(h);
  }, fadeOutMs);
}

function reveal(h: ActiveOverlay): void {
  h.revealTimer = null;
  h.revealDeadline = null;
  if (h.refs <= 0 || h.revealed || !h.driver) return;
  h.revealed = true;
  h.revealedAt = Date.now();
  // 文案取当前 driver 的（见 OverlayDriver 注释）：定时器可能由一个已
  // close 的并发调用方挂起（其文案已随 driver 栈回退丢弃），它只负责
  // "到点揭开"，不决定内容。
  render(h.dom, h.driver.state.reveal(), h.driver.text);
  if (!h.dom.root.isConnected) h.container.appendChild(h.dom.root);
  // 先落 opacity:0 再强制样式计算，再置 1——不这样做首帧就是终值，过渡不生效。
  h.dom.root.style.opacity = "0";
  void h.dom.root.offsetHeight;
  h.dom.root.style.opacity = "1";
}

const NOOP_HANDLE: ProgressOverlayHandle = {
  setProgress: () => {},
  setPhase: () => {},
  close: () => {},
};

function resolveTheme(
  theme: ProgressOverlayOptions["theme"],
): "light" | "dark" {
  if (theme === "light" || theme === "dark") return theme;
  return typeof window !== "undefined" &&
    typeof window.matchMedia === "function" &&
    window.matchMedia("(prefers-color-scheme: dark)").matches
    ? "dark"
    : "light";
}

/**
 * 显示全局进度遮罩，返回驱动它的句柄。
 *
 * 典型接线（以一个异步任务为例）：
 *
 * ```ts
 * const overlay = showProgressOverlay({
 *   text: { title: "正在导出", phases: { building: "构建中…", done: "收尾…" } },
 * });
 * try {
 *   await doWork({
 *     onProgress: (p) => overlay.setProgress(p),
 *     onStage: (key) => overlay.setPhase(key),
 *   });
 * } finally {
 *   overlay.close(); // 幂等
 * }
 * ```
 *
 * Node / SSR（无 `document`）下返回全空实现的句柄，调用方无需分支。
 */
export function showProgressOverlay(
  options: ProgressOverlayOptions = {},
): ProgressOverlayHandle {
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
      text: {
        title: options.text?.title ?? DEFAULT_TEXT.title,
        initial: options.text?.initial ?? DEFAULT_TEXT.initial,
        // phases 是嵌套对象，浅合并会整体替换掉默认（当前为空表，但保持
        // 合并语义一致，未来加默认项时不会静默丢失调用方的表）。
        phases: { ...DEFAULT_TEXT.phases, ...options.text?.phases },
        hint: options.text?.hint ?? DEFAULT_TEXT.hint,
      },
    };
    if (!cfg.container) return NOOP_HANDLE;
  } catch {
    return NOOP_HANDLE;
  }

  const state = new OverlayState({
    delayMs: cfg.delayMs,
    minVisibleMs: cfg.minVisibleMs,
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
        driver: null,
        driverStack: [],
        revealDeadline: null,
        revealTimer: null,
        unmountTimer: null,
        fadeTimer: null,
        revealed: false,
        revealedAt: null,
      };
    }
    const h = active;
    // 上一轮挂起的移除（淡出 / 最短可见延时）取消：本次复用同一份 DOM。
    clearTimer(h.unmountTimer);
    h.unmountTimer = null;
    // teardown 已把 opacity 压到 "0"、摘除只差 fadeTimer 到点——取消淡出后，
    // 复用节点要在下方分支里恢复 "1" 并重渲染，否则新任务的遮罩会以
    // opacity:0 挂完整场（不可见且仍在拦截交互）。
    clearTimer(h.fadeTimer);
    h.fadeTimer = null;
    // refs 归零后的复用是全新生命周期：栈中理论上已空（每次 close 都摘除
    // 自己），再清一次兜底上一场可能的异常残留。
    if (h.refs === 0) h.driverStack.length = 0;
    const driver: OverlayDriver = { state, text: cfg.text };
    h.driverStack.push(driver);
    h.driver = driver;
    h.refs += 1;

    if (!h.revealed) {
      // 揭示截止取并发调用方中的最晚者：任一活动调用方的 delayMs 未过都
      // 不揭示——后来者不会被前任已挂的定时器提前揭开，前任时序也不变。
      const deadline = Date.now() + state.delayRemaining();
      const merged =
        h.revealDeadline === null
          ? deadline
          : Math.max(h.revealDeadline, deadline);
      clearTimer(h.revealTimer);
      h.revealTimer = null;
      const wait = Math.max(0, merged - Date.now());
      if (wait === 0) {
        renderSafely(() => reveal(h));
      } else {
        h.revealDeadline = merged;
        // 定时器回调里的渲染异常无人接住（会以 uncaught 上报），同样走兜底。
        h.revealTimer = setTimeout(() => {
          renderSafely(() => reveal(h));
        }, wait);
      }
    } else {
      // 已揭示：并发接管显示中的遮罩（refs>1），或淡出/最短可见延迟窗口内
      // 的复活。立即按本次调用方重渲染——从接管的瞬间起，title/hint/label
      // 不得停留于前任调用方（此前只在淡出复活分支渲染，并发接管分支漏掉，
      // 文案会停留到新调用方收到第一个有效进度/阶段事件为止）。
      h.dom.root.style.opacity = "1";
      renderSafely(() => render(h.dom, state.snapshot(), cfg.text));
    }

    return {
      setProgress: (progress: number): void => {
        // 遮罩的任何异常都不允许影响宿主任务：整块吞掉。
        try {
          if (closed || active !== h) return;
          h.driver = driver;
          const snap = state.progress(progress);
          if (snap && h.revealed) render(h.dom, snap, cfg.text);
        } catch {
          /* 渲染失败不影响宿主任务 */
        }
      },
      setPhase: (key: string): void => {
        try {
          if (closed || active !== h) return;
          h.driver = driver;
          const snap = state.phase(key);
          if (snap && h.revealed) render(h.dom, snap, cfg.text);
        } catch {
          /* 渲染失败不影响宿主任务 */
        }
      },
      close: (): void => {
        try {
          if (closed) return;
          closed = true;
          if (active !== h) return;
          h.refs -= 1;
          // 从调用方栈摘除自己（末位与否都摘）：并发早退时 driver 才能回退。
          const stackIndex = h.driverStack.indexOf(driver);
          if (stackIndex >= 0) h.driverStack.splice(stackIndex, 1);
          // state.close() 只取其置 closed 的副作用（拒绝本 state 的后续事件），
          // 返回的 wait 不再使用——minVisible 改按 DOM 揭示时刻（h.revealedAt）
          // 计算：并发下 reveal 记在揭示当时的 driver 的 state 上，末位 close
          // 者直取自己的 state 会得出 0 而立即拆台，minVisible 形同虚设。
          state.close();
          if (h.refs > 0) {
            // 并发早退：driver 回退到最后介入的仍活动调用方。挂起的揭示定时
            // 器可能正是本调用方发起的，到点揭开时必须显示仍活动者的文案，
            // 而不是已结束的本调用方（否则遮罩会以"已结束任务"的文案挂到
            // 仍活动任务的全程）。
            h.driver = h.driverStack[h.driverStack.length - 1] ?? null;
            return;
          }
          // 关键：从未显示过时必须清掉延迟定时器，否则它会在 close 之后触发
          // 并把遮罩弹出来（任务比 delayMs 还快时必然发生）。
          clearTimer(h.revealTimer);
          h.revealTimer = null;
          h.revealDeadline = null;
          if (!h.revealed) {
            // 压根没插进 DOM，无需淡出
            h.driver = null;
            detach(h);
            return;
          }
          const elapsed =
            h.revealedAt === null ? Infinity : Date.now() - h.revealedAt;
          const wait = Math.max(0, cfg.minVisibleMs - elapsed);
          if (wait <= 0) teardown(h, cfg.fadeOutMs);
          else
            h.unmountTimer = setTimeout(() => {
              h.unmountTimer = null;
              teardown(h, cfg.fadeOutMs);
            }, wait);
        } catch {
          /* 关闭失败不影响宿主任务 */
        }
      },
    };
  } catch {
    return NOOP_HANDLE;
  }
}

/**
 * 让出一帧（实际两帧，确保样式与布局都已提交）。
 *
 * RAF 在不可见页面（后台标签、最小化窗口）中被浏览器完全暂停：没有兜底的
 * 话，后台触发的任务（定时报表等）会在本 Promise 上永久挂起，finally 里的
 * close 不执行，遮罩一旦揭开就永不结束。与一个短定时器竞速——等不到绘制就
 * 继续往下走，可见页面下双 RAF 几乎总是先到，行为不变。
 *
 * 供集成方使用：`showProgressOverlay` 后、同步阻塞开始前 await 它，是遮罩
 * 能被绘制出来的保证（见文件头注释的主线程阻塞说明）。
 */
export function nextPaint(): Promise<void> {
  if (typeof requestAnimationFrame !== "function") return Promise.resolve();
  return new Promise((resolve) => {
    let settled = false;
    const settle = () => {
      if (settled) return;
      settled = true;
      resolve();
    };
    requestAnimationFrame(() => requestAnimationFrame(settle));
    setTimeout(settle, 250);
  });
}
