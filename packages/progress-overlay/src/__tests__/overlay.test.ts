// @vitest-environment happy-dom
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { showProgressOverlay } from "../overlay";

/**
 * DOM 层测试。状态机自身的边界（延迟/最短可见/确定态切换）已在
 * overlay-state.test.ts 用可注入时钟精确覆盖，这里只验证容易出错的
 * DOM 生命周期：挂载时机、悬挂定时器、幂等关闭、无障碍属性、文案解析。
 */

const handles: { close: () => void }[] = [];

/** 默认关掉三个计时器：让挂载/移除在断言里都是同步可见的。 */
function show(opts: Parameters<typeof showProgressOverlay>[0] = {}) {
  const h = showProgressOverlay({
    delayMs: 0,
    minVisibleMs: 0,
    fadeOutMs: 0,
    ...opts,
  });
  handles.push(h);
  return h;
}

function root(): HTMLElement | null {
  return document.querySelector<HTMLElement>(".mxe-overlay");
}

beforeEach(() => {
  vi.useFakeTimers();
  document.body.innerHTML = "";
  document.head.innerHTML = "";
});

afterEach(() => {
  // 模块级单例跨用例存活，必须逐个关干净并跑完淡出，否则下一个用例
  // 会复用到上一个用例残留的 DOM 与 revealed 状态。
  for (const h of handles) h.close();
  handles.length = 0;
  vi.advanceTimersByTime(10_000);
  vi.useRealTimers();
});

describe("遮罩生命周期", () => {
  it("delayMs 为 0 时同步挂载并完成入场过渡", () => {
    show();
    const el = root();
    expect(el).not.toBeNull();
    expect(el!.parentElement).toBe(document.body);
    expect(el!.style.opacity).toBe("1");
  });

  it("延迟未到不挂载，到点才挂载", () => {
    show({ delayMs: 200 });
    expect(root()).toBeNull();
    vi.advanceTimersByTime(199);
    expect(root()).toBeNull();
    vi.advanceTimersByTime(1);
    expect(root()).not.toBeNull();
  });

  it("任务快于 delayMs 时遮罩永不出现——悬挂的延迟定时器必须被清掉", () => {
    const h = show({ delayMs: 200 });
    h.close();
    // 迟到的定时器若没被清掉，会在这里把遮罩弹出来
    vi.advanceTimersByTime(1_000);
    expect(root()).toBeNull();
  });

  it("close 幂等，重复调用不再产生副作用", () => {
    const h = show();
    h.close();
    h.close();
    vi.advanceTimersByTime(1_000);
    expect(root()).toBeNull();
  });

  it("最短可见时长未满时不移除，到时先淡出再移除", () => {
    const h = show({ minVisibleMs: 300, fadeOutMs: 100 });
    h.close();
    vi.advanceTimersByTime(299);
    expect(root()).not.toBeNull();
    vi.advanceTimersByTime(1); // 最短可见到时 -> 开始淡出
    expect(root()).not.toBeNull();
    expect(root()!.style.opacity).toBe("0");
    vi.advanceTimersByTime(100); // 淡出结束 -> 移除
    expect(root()).toBeNull();
  });

  it("并发任务共用一份 DOM，最后一个关闭才移除", () => {
    const a = show({ minVisibleMs: 0 });
    const b = show({ minVisibleMs: 0 });
    expect(document.querySelectorAll(".mxe-overlay")).toHaveLength(1);
    a.close();
    vi.advanceTimersByTime(1_000);
    expect(root()).not.toBeNull();
    b.close();
    vi.advanceTimersByTime(1_000);
    expect(root()).toBeNull();
  });

  it("淡出进行中被新任务复用时恢复可见（回归：opacity 曾停在 0）", () => {
    // 第一次任务关闭后开始淡出：teardown 已把 opacity 压到 "0"
    const a = show({ minVisibleMs: 0, fadeOutMs: 100 });
    a.close();
    expect(root()!.style.opacity).toBe("0");

    // 淡出窗口内发起第二次任务：复用同一节点，必须恢复可见并回到初始文案
    const b = show();
    expect(root()!.style.opacity).toBe("1");
    expect(root()!.querySelector(".mxe-label")!.textContent).toBe("正在处理…");
    expect(root()!.dataset.mxeMode).toBe("indeterminate");

    // 第一次任务的淡出定时器已被取消，到点不得摘除正在使用的节点
    vi.advanceTimersByTime(200);
    expect(root()).not.toBeNull();

    b.close();
    vi.advanceTimersByTime(1_000);
    expect(root()).toBeNull();
  });

  it("并发下延迟定时器揭开遮罩时，文案取当前驱动者的（回归：曾泄漏发起者的文案）", () => {
    // A 先 show：refs=1，挂 200ms 的 revealTimer（闭包携带 A 的文案）
    const a = show({ delayMs: 200, text: { title: "A 的任务" } });
    // B 再 show：refs=2，与已挂定时器的截止取更晚者（B 的 delayMs=0 不顺延），
    // driver 换成 B
    const b = show({ text: { title: "B 的任务" } });
    expect(root()).toBeNull(); // A 的 delayMs 未到，遮罩未显示

    a.close(); // refs 2->1：close 提前返回，A 的 revealTimer 仍然存活
    vi.advanceTimersByTime(200); // 定时器到点揭开——内容必须属于 B
    expect(root()).not.toBeNull();
    expect(root()!.querySelector(".mxe-title")!.textContent).toBe("B 的任务");

    b.close();
    vi.advanceTimersByTime(1_000);
    expect(root()).toBeNull();
  });

  it("并发接管已显示的遮罩时立即重渲染文案（回归：曾停留前任文案直到新调用方首个有效事件）", () => {
    const a = show({ text: { title: "A 的任务" } });
    expect(root()!.querySelector(".mxe-title")!.textContent).toBe("A 的任务");
    // A 已揭示且仍在任务中，B 接管（refs=2）：接管分支此前不渲染，title/
    // hint/label 会一直显示 A 的，直到 B 收到第一个有效进度/阶段事件。
    const b = show({ text: { title: "B 的任务" } });
    expect(root()!.querySelector(".mxe-title")!.textContent).toBe("B 的任务");
    expect(root()!.querySelector(".mxe-label")!.textContent).toBe("正在处理…");

    a.close();
    b.close();
    vi.advanceTimersByTime(1_000);
    expect(root()).toBeNull();
  });

  it("并发后来者更晚的 delayMs 顺延揭示（回归：揭示曾被前任的定时器提前）", () => {
    const a = show({ delayMs: 200 });
    // B 要求 2000ms 内不打扰：揭示截止应取并发调用方中的最晚者
    const b = show({ delayMs: 2_000 });
    vi.advanceTimersByTime(1_999);
    expect(root()).toBeNull(); // A 的 200ms 定时器到点也不得提前揭示
    vi.advanceTimersByTime(1);
    expect(root()).not.toBeNull();

    a.close();
    b.close();
    vi.advanceTimersByTime(1_000);
    expect(root()).toBeNull();
  });

  it("早退调用方挂起的揭示定时器到点时，文案回退到仍活动调用方（回归：曾显示已结束调用方的文案）", () => {
    // A 挂起 200ms 定时器；B show 后重挂为两者中更晚的截止（遮罩级定时器）
    const a = show({ delayMs: 200, text: { title: "A 的任务" } });
    const b = show({ text: { title: "B 的任务" } });
    // B 先结束（refs 2->1 早退）：driver 必须按栈回退到 A，而非留在已结束的 B
    b.close();
    vi.advanceTimersByTime(200);
    expect(root()).not.toBeNull();
    expect(root()!.querySelector(".mxe-title")!.textContent).toBe("A 的任务");

    a.close();
    vi.advanceTimersByTime(1_000);
    expect(root()).toBeNull();
  });

  it("并发揭示后末位 close 的最短可见按 DOM 揭示时刻计（回归：曾立即拆台一闪而过）", () => {
    // A 挂定时器期间 B 接管并成为揭示时的 driver；B 先结束，A 末位 close 时
    // 自己的 state 从未 reveal——修复前 state.close() 算出 wait=0 立即淡出。
    const a = show({ delayMs: 200, minVisibleMs: 300, fadeOutMs: 100 });
    const b = show({ minVisibleMs: 300, fadeOutMs: 100 });
    vi.advanceTimersByTime(200); // 揭示：内容 driver=B，A 的 state 未记 revealedAt
    expect(root()).not.toBeNull();
    b.close();
    a.close(); // 末位 close：仍须遵守 minVisible（按 h.revealedAt 而非 state_A）
    expect(root()!.style.opacity).toBe("1"); // 不得立即开始淡出
    vi.advanceTimersByTime(300);
    expect(root()!.style.opacity).toBe("0"); // 最短可见到时才开始淡出
    vi.advanceTimersByTime(100);
    expect(root()).toBeNull();
  });
});

describe("遮罩渲染", () => {
  it("中间进度切确定态并同步宽度、aria-valuenow 与百分比", () => {
    const h = show();
    const el = root()!;
    const bar = el.querySelector<HTMLElement>(".mxe-bar")!;
    expect(el.dataset.mxeMode).toBe("indeterminate");
    expect(bar.hasAttribute("aria-valuenow")).toBe(false);

    h.setProgress(0.42);
    expect(el.dataset.mxeMode).toBe("determinate");
    expect(el.querySelector<HTMLElement>(".mxe-fill")!.style.transform).toBe(
      "scaleX(0.42)",
    );
    expect(bar.getAttribute("aria-valuenow")).toBe("42");
    expect(el.querySelector<HTMLElement>(".mxe-percent")!.textContent).toBe(
      "42%",
    );
  });

  it("不确定态渲染 svg 圆环（对读屏纯装饰），百分比节点初始为空", () => {
    show();
    const el = root()!;
    const spinner = el.querySelector<SVGElement>(".mxe-spinner")!;
    expect(spinner.tagName.toLowerCase()).toBe("svg");
    expect(spinner.getAttribute("aria-hidden")).toBe("true");
    expect(spinner.querySelector(".mxe-spinner-arc")).not.toBeNull();
    // 百分比独立于 label 节点（aria-live 区域不混入进度数字）
    expect(el.querySelector<HTMLElement>(".mxe-percent")!.textContent).toBe("");
  });

  it("入口的 0 与收尾的 1 都不改变不确定态（无中间进度的调用方）", () => {
    const h = show();
    h.setProgress(0);
    h.setProgress(1);
    expect(root()!.dataset.mxeMode).toBe("indeterminate");
  });

  it("阶段 key 查 phases 表切文案；未覆盖的 key 原样显示", () => {
    const h = show({
      text: {
        initial: "准备中…",
        phases: { building: "构建中…", downloading: "下载中…" },
      },
    });
    // 写成块体而非 `root()!.querySelector(...)!` 链：后者会被
    // no-unnecessary-type-assertion 误报（去掉 ! 后 tsc 仍报 TS2531）
    const label = (): string => {
      const el = root();
      return el?.querySelector(".mxe-label")?.textContent ?? "";
    };
    expect(label()).toBe("准备中…");
    h.setPhase("building");
    expect(label()).toBe("构建中…");
    h.setPhase("downloading");
    expect(label()).toBe("下载中…");
    // 未在 phases 声明的 key：原样显示（可发现拼写错误，而非静默停留旧文案）
    h.setPhase("uploading");
    expect(label()).toBe("uploading");
  });

  it("blockInteraction 默认阻断交互（不放开 pointer-events）", () => {
    show();
    expect(root()!.style.pointerEvents).toBe("");
  });

  it("blockInteraction: false 时只做视觉覆盖", () => {
    show({ blockInteraction: false });
    expect(root()!.style.pointerEvents).toBe("none");
  });

  it("按主题挂到 root 的 data 属性上（auto 解析为 light/dark）", () => {
    show({ theme: "dark" });
    expect(root()!.dataset.mxeTheme).toBe("dark");
  });
});

describe("非浏览器环境", () => {
  it("无 document 时返回空实现句柄，调用不抛错", () => {
    vi.stubGlobal("document", undefined);
    const h = showProgressOverlay();
    expect(() => {
      h.setProgress(0.5);
      h.setPhase("building");
      h.close();
    }).not.toThrow();
  });
});
