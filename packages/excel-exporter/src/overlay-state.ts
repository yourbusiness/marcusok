import type { ExportPhase } from "./types";

/**
 * 遮罩的纯状态机：不碰 DOM、不持有定时器，只做决策，由 overlay.ts 驱动。
 * 拆出来的目的是让「延迟门控 / 最短可见 / 确定态切换 / 文案时序 / 关闭幂等」
 * 这些最容易写错的边界能在 node 环境下直接单测（仓库默认测试环境无 DOM）。
 */

/**
 * 阶段文案 key。`onPhase` 只在阶段**结束后**回调（见 types.ts），没有
 * "阶段开始"事件，所以文案只能乐观推进：收到 `init` 即表示 init 已完、当前
 * 在做 build，文案切 `building`，依此类推。
 */
export type OverlayLabel =
  "preparing" | "building" | "downloading" | "finishing";

/** 可渲染状态快照。 */
export interface OverlaySnapshot {
  /** null = 不确定态；0..1 = 确定态（渲染为确定宽度）。 */
  determinate: number | null;
  label: OverlayLabel;
}

export interface OverlayStateConfig {
  /** 显示前延迟（ms）：导出在此之前结束则遮罩完全不出现。 */
  delayMs: number;
  /** 最短可见时长（ms）：已显示后关闭时不足此值则延后移除，避免一闪而过。 */
  minVisibleMs: number;
  /** 本路由是否有 download 阶段，决定 build 之后切哪个文案。 */
  willDownload: boolean;
}

export class OverlayState {
  private readonly cfg: OverlayStateConfig;
  private readonly now: () => number;
  private readonly startedAt: number;
  private revealedAt: number | null = null;
  private closed = false;
  private determinate: number | null = null;
  private label: OverlayLabel = "preparing";

  constructor(cfg: OverlayStateConfig, now: () => number = () => Date.now()) {
    this.cfg = cfg;
    this.now = now;
    this.startedAt = now();
  }

  /** 是否已经显示过。关闭时据此决定「淡出移除」还是「直接丢弃」。 */
  get isRevealed(): boolean {
    return this.revealedAt !== null;
  }

  /** 是否已请求关闭。延迟定时器到期后据此放弃显示。 */
  get isClosed(): boolean {
    return this.closed;
  }

  /** 距允许显示还剩多少毫秒；0 表示可以立即显示。 */
  delayRemaining(): number {
    return Math.max(0, this.cfg.delayMs - (this.now() - this.startedAt));
  }

  /** 延迟到期，允许显示。已关闭时为 no-op（迟到的一帧不得让遮罩复活）。 */
  reveal(): OverlaySnapshot {
    if (!this.closed && this.revealedAt === null) this.revealedAt = this.now();
    return this.snapshot();
  }

  /**
   * 收到进度。返回需要重渲染的新快照，无变化返回 null。
   *
   * 只有 stream 路径会产生 0..1 之间的中间值（每 1000 行一次）；其余路由
   * （main/worker 的 workbook 分支、worker 首次之后的导出）只有入口的 0 与
   * 收尾的 1，因此它们会全程停留在不确定态——这是数据源本身的粒度，不是
   * 渲染问题。
   */
  progress(p: number): OverlaySnapshot | null {
    if (this.closed) return null;
    // 入口的 0（以及 NaN 之类的脏值）不改变状态：0 只说明导出已开始。
    if (!(p > 0)) return null;
    if (p < 1) {
      if (this.determinate === p) return null;
      this.determinate = p;
      return this.snapshot();
    }
    // p >= 1 不关闭遮罩：失败路径同样会发收尾的 1（types.ts 契约），成功后
    // 还有 download 要跑，关闭时机由调用方的 finally 掌握。只有已经是确定态
    // 时才推到 100%，避免整个导出过程中未曾有过真实进度的场景（workbook 路由）
    // 凭这一个 1 假装走完全程。
    if (this.determinate === null || this.determinate === 1) return null;
    this.determinate = 1;
    return this.snapshot();
  }

  /** 收到阶段完成事件。返回需要重渲染的新快照，无变化返回 null。 */
  phase(phase: ExportPhase): OverlaySnapshot | null {
    if (this.closed) return null;
    const next: OverlayLabel =
      phase === "init"
        ? "building"
        : phase === "build"
          ? this.cfg.willDownload
            ? "downloading"
            : "finishing"
          : "finishing";
    if (next === this.label) return null;
    this.label = next;
    return this.snapshot();
  }

  /**
   * 请求关闭，返回移除 DOM 前还需等待的毫秒数。
   *
   * 从未显示过 -> 0：无需移除，但调用方**必须**同时清掉延迟定时器，否则
   * 导出比 delayMs 还快时，悬挂的定时器会在导出结束后把遮罩弹出来。
   * 已显示但不足 minVisibleMs -> 剩余毫秒；否则 0。
   */
  close(): number {
    if (this.closed) return 0;
    this.closed = true;
    if (this.revealedAt === null) return 0;
    return Math.max(0, this.cfg.minVisibleMs - (this.now() - this.revealedAt));
  }

  snapshot(): OverlaySnapshot {
    return { determinate: this.determinate, label: this.label };
  }
}
