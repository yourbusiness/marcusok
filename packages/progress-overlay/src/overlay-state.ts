/**
 * 遮罩的纯状态机：不碰 DOM、不持有定时器，只做决策，由 overlay.ts 驱动。
 * 拆出来的目的是让「延迟门控 / 最短可见 / 确定态切换 / 文案时序 / 关闭幂等」
 * 这些最容易写错的边界能在 node 环境下直接单测（仓库默认测试环境无 DOM）。
 */

/** 可渲染状态快照。 */
export interface OverlaySnapshot {
  /** null = 不确定态；0..1 = 确定态（渲染为确定宽度与百分比）。 */
  determinate: number | null;
  /**
   * null = 初始（渲染 `text.initial`）；否则为调用方定义的阶段 key，渲染层
   * 查 `text.phases` 得文案。状态机不解释 key 的含义——"哪个事件之后显示
   * 哪段文案"的映射是业务侧（如 excel-exporter 的 ExportPhase 推进）的事。
   */
  phase: string | null;
}

export interface OverlayStateConfig {
  /** 显示前延迟（ms）：任务在此之前结束则遮罩完全不出现。 */
  delayMs: number;
  /** 最短可见时长（ms）：已显示后关闭时不足此值则延后移除，避免一闪而过。 */
  minVisibleMs: number;
}

export class OverlayState {
  private readonly cfg: OverlayStateConfig;
  private readonly now: () => number;
  private readonly startedAt: number;
  private revealedAt: number | null = null;
  private closed = false;
  private determinate: number | null = null;
  private phaseKey: string | null = null;

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
   * 入口的 0（以及 NaN 之类的脏值）不改变状态：0 只说明任务已开始；中间值
   * (0,1) 一到即切确定态——不确定态（spinner）的时长完全由数据源粒度决定，
   * 全程没有中间进度的调用方会一直停在不确定态，这是预期而非渲染问题。
   */
  progress(p: number): OverlaySnapshot | null {
    if (this.closed) return null;
    if (!(p > 0)) return null;
    if (p < 1) {
      if (this.determinate === p) return null;
      this.determinate = p;
      return this.snapshot();
    }
    // p >= 1 不触发关闭：失败路径同样可能发收尾的 1，关闭时机由调用方的
    // finally 掌握。只有已经是确定态时才推到 100%，避免整个过程中未曾有
    // 过真实进度的场景凭这一个 1 假装走完全程。
    if (this.determinate === null || this.determinate === 1) return null;
    this.determinate = 1;
    return this.snapshot();
  }

  /**
   * 收到阶段事件。返回需要重渲染的新快照，无变化返回 null。
   * key 任意（由调用方定义）；与上次相同则不重渲染（label 是 aria-live
   * 区域，重复写同一文本会被读屏重复播报）。
   */
  phase(key: string): OverlaySnapshot | null {
    if (this.closed) return null;
    if (this.phaseKey === key) return null;
    this.phaseKey = key;
    return this.snapshot();
  }

  /**
   * 请求关闭，返回移除 DOM 前还需等待的毫秒数。
   *
   * 从未显示过 -> 0：无需移除，但调用方**必须**同时清掉延迟定时器，否则
   * 任务比 delayMs 还快时，悬挂的定时器会在任务结束后把遮罩弹出来。
   * 已显示但不足 minVisibleMs -> 剩余毫秒；否则 0。
   */
  close(): number {
    if (this.closed) return 0;
    this.closed = true;
    if (this.revealedAt === null) return 0;
    return Math.max(0, this.cfg.minVisibleMs - (this.now() - this.revealedAt));
  }

  snapshot(): OverlaySnapshot {
    return { determinate: this.determinate, phase: this.phaseKey };
  }
}
