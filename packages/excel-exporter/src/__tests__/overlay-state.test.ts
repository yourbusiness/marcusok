import { describe, it, expect } from "vitest";
import { OverlayState, type OverlayStateConfig } from "../overlay-state";

/** 可注入时钟：状态机的延迟/最短可见判定全部走 now()，单测据此精确控时。 */
function makeClock(start = 1_000) {
  let t = start;
  return {
    now: () => t,
    advance: (ms: number) => {
      t += ms;
    },
  };
}

function makeState(over: Partial<OverlayStateConfig> = {}) {
  const clock = makeClock();
  const state = new OverlayState(
    { delayMs: 200, minVisibleMs: 300, willDownload: true, ...over },
    clock.now,
  );
  return { state, clock };
}

describe("OverlayState 延迟门控", () => {
  it("初始为 preparing 且是不确定态", () => {
    const { state } = makeState();
    expect(state.snapshot()).toEqual({
      determinate: null,
      label: "preparing",
    });
    expect(state.isRevealed).toBe(false);
  });

  it("delayRemaining 随时钟递减，到点后归零且不再为负", () => {
    const { state, clock } = makeState({ delayMs: 200 });
    expect(state.delayRemaining()).toBe(200);
    clock.advance(120);
    expect(state.delayRemaining()).toBe(80);
    clock.advance(500);
    expect(state.delayRemaining()).toBe(0);
  });

  it("reveal 会记录显示时刻（影响最短可见判定）", () => {
    const { state, clock } = makeState();
    expect(state.isRevealed).toBe(false);
    clock.advance(200);
    state.reveal();
    expect(state.isRevealed).toBe(true);
  });

  it("已关闭后再 reveal 是 no-op，遮罩不得复活", () => {
    const { state } = makeState();
    state.close();
    state.reveal();
    expect(state.isRevealed).toBe(false);
  });
});

describe("OverlayState 关闭语义", () => {
  it("从未显示就关闭 -> 0，调用方据此直接丢弃并清掉延迟定时器", () => {
    const { state, clock } = makeState();
    clock.advance(50);
    expect(state.close()).toBe(0);
    expect(state.isClosed).toBe(true);
  });

  it("已显示但不足 minVisibleMs -> 返回剩余等待毫秒", () => {
    const { state, clock } = makeState({ minVisibleMs: 300 });
    state.reveal();
    clock.advance(100);
    expect(state.close()).toBe(200);
  });

  it("已显示且超过 minVisibleMs -> 立即移除", () => {
    const { state, clock } = makeState({ minVisibleMs: 300 });
    state.reveal();
    clock.advance(1000);
    expect(state.close()).toBe(0);
  });

  it("close 幂等：第二次调用不再产生等待", () => {
    const { state, clock } = makeState({ minVisibleMs: 300 });
    state.reveal();
    clock.advance(10);
    expect(state.close()).toBe(290);
    expect(state.close()).toBe(0);
  });

  it("关闭后进度与阶段事件都被忽略", () => {
    const { state } = makeState();
    state.close();
    expect(state.progress(0.5)).toBeNull();
    expect(state.phase("build")).toBeNull();
  });
});

describe("OverlayState 进度", () => {
  it("入口的 0（及 NaN）不改变状态", () => {
    const { state } = makeState();
    expect(state.progress(0)).toBeNull();
    expect(state.progress(Number.NaN)).toBeNull();
    expect(state.snapshot().determinate).toBeNull();
  });

  it("0..1 之间的中间值切到确定态", () => {
    const { state } = makeState();
    expect(state.progress(0.42)).toEqual({
      determinate: 0.42,
      label: "preparing",
    });
    expect(state.progress(0.43)?.determinate).toBe(0.43);
  });

  it("重复上报同一进度值不产生新快照", () => {
    const { state } = makeState();
    state.progress(0.5);
    expect(state.progress(0.5)).toBeNull();
  });

  it("收尾的 1 不关闭遮罩，仅在已是确定态时推到 100%", () => {
    const { state } = makeState();
    // 全程没有中间进度（workbook 路由）：不凭这一个 1 假装走完全程
    expect(state.progress(1)).toBeNull();
    expect(state.snapshot().determinate).toBeNull();
  });

  it("已有真实进度时，收尾的 1 把确定态推到 100%", () => {
    const { state } = makeState();
    state.progress(0.87);
    expect(state.progress(1)?.determinate).toBe(1);
  });
});

describe("OverlayState 阶段文案", () => {
  it("init 完成 -> 构建中（onPhase 只在阶段结束后回调，文案只能乐观推进）", () => {
    const { state } = makeState();
    expect(state.phase("init")?.label).toBe("building");
  });

  it("build 完成 -> 有 download 时切下载中，否则切即将完成", () => {
    const withDownload = makeState({ willDownload: true });
    expect(withDownload.state.phase("build")?.label).toBe("downloading");

    const noDownload = makeState({ willDownload: false });
    expect(noDownload.state.phase("build")?.label).toBe("finishing");
  });

  it("download 完成 -> 即将完成", () => {
    const { state } = makeState();
    state.phase("build");
    expect(state.phase("download")?.label).toBe("finishing");
  });

  it("文案无变化时不返回新快照", () => {
    const { state } = makeState();
    state.phase("init");
    expect(state.phase("init")).toBeNull();
  });

  it("阶段推进不影响已经拿到的确定态", () => {
    const { state } = makeState();
    state.progress(0.6);
    expect(state.phase("build")).toEqual({
      determinate: 0.6,
      label: "downloading",
    });
  });
});
