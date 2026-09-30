// @vitest-environment happy-dom
import { describe, it, expect, vi, beforeEach } from "vitest";
import { exportExcel } from "../index";
import type { ProgressOverlayHandle } from "../index";
import type { ExportOptions, SheetConfig } from "../types";

/**
 * 主入口 overlay 接线测试：默认开启 / overlay:false 关闭 / 定制合并 / 失败
 * 关闭。遮罩本体（DOM 生命周期、并发、无障碍）在 @marcusok/progress-overlay
 * 包内覆盖，这里 mock 掉该包，只验证 exportExcel 包装层把它接对了。
 *
 * 导出走显式 stream + 单行数据：纯 JS 路径不加载 WASM，happy-dom 下可跑。
 */

const calls: string[] = [];
let handle: ProgressOverlayHandle;
let showOptions: Record<string, unknown> | null = null;
let nextPaintCalls = 0;

vi.mock("@marcusok/progress-overlay", () => ({
  showProgressOverlay: vi.fn((options: Record<string, unknown>) => {
    showOptions = options;
    return handle;
  }),
  nextPaint: vi.fn(() => {
    nextPaintCalls += 1;
    return Promise.resolve();
  }),
}));

beforeEach(() => {
  calls.length = 0;
  showOptions = null;
  nextPaintCalls = 0;
  handle = {
    setProgress: (p: number) => calls.push(`progress:${p}`),
    setPhase: (key: string) => calls.push(`phase:${key}`),
    close: () => calls.push("close"),
  };
});

const SHEET: SheetConfig = {
  name: "S",
  columns: [{ prop: "a", label: "A" }],
  data: [{ a: 1 }],
};

/** 单行 + 显式 stream：绕开 WASM 与 Worker，happy-dom 下直跑纯 JS 路径。 */
function baseOptions(overlay?: ExportOptions["overlay"]): ExportOptions {
  return {
    filename: "overlay-wiring.xlsx",
    mode: "stream",
    sheets: [SHEET],
    download: false,
    ...(overlay !== undefined ? { overlay } : {}),
  };
}

describe("exportExcel overlay 接线", () => {
  it("默认（无 overlay 字段）显示遮罩：挂载→让帧→链式回调→finally 关闭", async () => {
    const seen: number[] = [];
    await exportExcel({
      ...baseOptions(),
      onProgress: (p) => seen.push(p),
    });

    expect(showOptions).not.toBeNull();
    expect(nextPaintCalls).toBe(1);
    // 链式追加：调用方回调照常收 0 -> ... -> 1
    expect(seen[0]).toBe(0);
    expect(seen[seen.length - 1]).toBe(1);
    // 遮罩句柄收到同样的进度序列，并最终被 close（且只 close 一次）
    expect(calls).toContain("progress:0");
    expect(calls[calls.length - 1]).toBe("close");
    expect(calls.filter((c) => c === "close")).toHaveLength(1);
  });

  it("默认文案带导出语义（title 与 phases 表）", async () => {
    await exportExcel(baseOptions());
    const text = showOptions!.text as Record<string, unknown>;
    expect(text.title).toBe("正在导出 Excel");
    expect(text.initial).toBe("准备中…");
    expect((text.phases as Record<string, string>).building).toBe(
      "正在构建工作簿…",
    );
  });

  it("overlay: false 完全关闭：不挂遮罩、不让帧、回调原样", async () => {
    const seen: number[] = [];
    await exportExcel({
      ...baseOptions(false),
      onProgress: (p) => seen.push(p),
    });
    expect(showOptions).toBeNull();
    expect(nextPaintCalls).toBe(0);
    expect(calls).toHaveLength(0);
    expect(seen[0]).toBe(0);
    expect(seen[seen.length - 1]).toBe(1);
  });

  it("overlay: {...} 定制：title 覆盖，phases 与导出默认合并", async () => {
    await exportExcel(
      baseOptions({ delayMs: 0, text: { title: "自定义标题" } }),
    );
    const text = showOptions!.text as Record<string, unknown>;
    expect(text.title).toBe("自定义标题");
    // 未覆盖的 phases 项保留导出默认：setPhase("building") 不会显示裸 key
    expect((text.phases as Record<string, string>).building).toBe(
      "正在构建工作簿…",
    );
    expect(showOptions!.delayMs).toBe(0);
  });

  it("阶段事件都映射到导出语义 key（init→building 等乐观推进）", async () => {
    await exportExcel(baseOptions());
    const phases = calls.filter((c) => c.startsWith("phase:"));
    expect(phases.length).toBeGreaterThan(0);
    for (const p of phases) {
      expect([
        "phase:building",
        "phase:downloading",
        "phase:finishing",
      ]).toContain(p);
    }
  });

  it("导出失败（非法输入）时 close 仍被执行且只执行一次", async () => {
    const result = await exportExcel({
      ...baseOptions(),
      sheets: [],
    });
    expect(result.success).toBe(false);
    expect(calls.filter((c) => c === "close")).toHaveLength(1);
  });
});
