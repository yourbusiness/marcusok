// @vitest-environment happy-dom
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import type { ExportOptions } from "../types";
import { showExportOverlay, exportExcelWithOverlay } from "../overlay";

/**
 * DOM 层测试。状态机自身的边界（延迟/最短可见/确定态切换）已在
 * overlay-state.test.ts 用可注入时钟精确覆盖，这里只验证容易出错的
 * DOM 生命周期：挂载时机、悬挂定时器、幂等关闭、无障碍属性、以及
 * `exportExcelWithOverlay` 不吞掉调用方原有回调。
 */

// 引擎不必真的跑：overlay.ts 只从 ./index 取 exportExcel 这一个值绑定。
// 用全合成工厂替换，避免在 happy-dom 里加载 WASM 引擎。
const probe = vi.hoisted(() => ({
  overlayPresent: false,
  modeAfterProgress: "",
  labelAfterPhase: "",
}));

vi.mock("../index", () => ({
  exportExcel: vi.fn((options: ExportOptions) => {
    const root = document.querySelector<HTMLElement>(".mxe-overlay");
    probe.overlayPresent = root !== null;
    options.onProgress?.(0);
    options.onProgress?.(0.5);
    probe.modeAfterProgress = root?.dataset.mxeMode ?? "";
    options.onPhase?.("init", 1);
    probe.labelAfterPhase =
      root?.querySelector(".mxe-label")?.textContent ?? "";
    options.onProgress?.(1);
    return { success: true, rowCount: 0 };
  }),
}));

const handles: { close: () => void }[] = [];

/** 默认关掉三个计时器：让挂载/移除在断言里都是同步可见的。 */
function show(opts: Parameters<typeof showExportOverlay>[0] = {}) {
  const h = showExportOverlay({
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
  probe.overlayPresent = false;
  probe.modeAfterProgress = "";
  probe.labelAfterPhase = "";
});

afterEach(() => {
  // 模块级单例跨用例存活，必须逐个关干净并跑完淡出，否则下一个用例
  // 会复用到上一个用例残留的 DOM 与 revealed 状态。
  for (const h of handles) h.close();
  handles.length = 0;
  vi.advanceTimersByTime(10_000);
  vi.useRealTimers();
  vi.unstubAllGlobals();
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

  it("导出快于 delayMs 时遮罩永不出现——悬挂的延迟定时器必须被清掉", () => {
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

  it("并发导出共用一份 DOM，最后一个关闭才移除", () => {
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

  it("淡出进行中被新导出复用时恢复可见（回归：opacity 曾停在 0）", () => {
    // 第一次导出关闭后开始淡出：teardown 已把 opacity 压到 "0"
    const a = show({ minVisibleMs: 0, fadeOutMs: 100 });
    a.close();
    expect(root()!.style.opacity).toBe("0");

    // 淡出窗口内发起第二次导出：复用同一节点，必须恢复可见并回到初始文案
    const b = show();
    expect(root()!.style.opacity).toBe("1");
    expect(root()!.querySelector(".mxe-label")!.textContent).toBe("准备中…");
    expect(root()!.dataset.mxeMode).toBe("indeterminate");

    // 第一次导出的淡出定时器已被取消，到点不得摘除正在使用的节点
    vi.advanceTimersByTime(200);
    expect(root()).not.toBeNull();

    b.close();
    vi.advanceTimersByTime(1_000);
    expect(root()).toBeNull();
  });
});

describe("遮罩渲染", () => {
  it("中间进度切确定态并同步宽度与 aria-valuenow", () => {
    const h = show();
    const el = root()!;
    const bar = el.querySelector<HTMLElement>(".mxe-bar")!;
    expect(el.dataset.mxeMode).toBe("indeterminate");
    expect(bar.hasAttribute("aria-valuenow")).toBe(false);

    h.handleProgress(0.42);
    expect(el.dataset.mxeMode).toBe("determinate");
    expect(el.querySelector<HTMLElement>(".mxe-fill")!.style.transform).toBe(
      "scaleX(0.42)",
    );
    expect(bar.getAttribute("aria-valuenow")).toBe("42");
  });

  it("入口的 0 与收尾的 1 都不改变不确定态（workbook 路由无中间进度）", () => {
    const h = show();
    h.handleProgress(0);
    h.handleProgress(1);
    expect(root()!.dataset.mxeMode).toBe("indeterminate");
  });

  it("阶段完成推进文案：init -> 构建中，build -> 下载中", () => {
    const h = show();
    // 写成块体而非 `root()!.querySelector(...)!` 链：后者会被
    // no-unnecessary-type-assertion 误报（去掉 ! 后 tsc 仍报 TS2531）
    const label = (): string => {
      const el = root();
      return el?.querySelector(".mxe-label")?.textContent ?? "";
    };
    expect(label()).toBe("准备中…");
    h.handlePhase("init");
    expect(label()).toBe("正在构建工作簿…");
    h.handlePhase("build");
    expect(label()).toBe("正在下载…");
  });

  it("download: false 的路由在 build 后落在「即将完成」而不是「正在下载」", () => {
    const h = show({ willDownload: false });
    h.handlePhase("build");
    expect(root()!.querySelector(".mxe-label")!.textContent).toBe("即将完成…");
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

describe("exportExcelWithOverlay", () => {
  it("链式保留调用方已有的 onProgress / onPhase", async () => {
    const seen: number[] = [];
    const phases: string[] = [];
    const promise = exportExcelWithOverlay(
      {
        filename: "x.xlsx",
        sheets: [],
        onProgress: (p) => seen.push(p),
        onPhase: (phase) => phases.push(phase),
      },
      { delayMs: 0, minVisibleMs: 0, fadeOutMs: 0 },
    );
    // nextPaint 让出两帧后才真正调用 exportExcel，先把假定时器推过去
    await vi.advanceTimersByTimeAsync(100);
    const result = await promise;

    expect(result.success).toBe(true);
    expect(seen).toEqual([0, 0.5, 1]);
    expect(phases).toEqual(["init"]);
    // 导出进行中遮罩确实在 DOM 里，且被进度驱动到了确定态
    expect(probe.overlayPresent).toBe(true);
    expect(probe.modeAfterProgress).toBe("determinate");
    expect(probe.labelAfterPhase).toBe("正在构建工作簿…");
    // 导出结束后遮罩必须移除（成功路径由 finally 关闭）
    expect(root()).toBeNull();
  });

  it("导出抛错时遮罩同样被移除", async () => {
    const { exportExcel } = await import("../index");
    vi.mocked(exportExcel).mockRejectedValueOnce(new Error("boom"));
    const promise = exportExcelWithOverlay(
      { filename: "x.xlsx", sheets: [] },
      { delayMs: 0, minVisibleMs: 0, fadeOutMs: 0 },
    );
    // 同步挂上断言：否则拒绝会先于处理器出现，被 Node 判为 unhandled rejection
    const assertion = expect(promise).rejects.toThrow("boom");
    await vi.advanceTimersByTimeAsync(100);
    await assertion;
    expect(root()).toBeNull();
  });
});

describe("非浏览器环境", () => {
  it("无 document 时返回空实现句柄，调用不抛错", () => {
    vi.stubGlobal("document", undefined);
    const h = showExportOverlay();
    expect(() => {
      h.handleProgress(0.5);
      h.handlePhase("build");
      h.close();
    }).not.toThrow();
  });
});
