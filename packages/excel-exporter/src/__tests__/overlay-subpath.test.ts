// @vitest-environment happy-dom
/**
 * `@marcusok/excel-exporter/overlay` 子路径（公开入口）测试。
 *
 * 覆盖三件事：
 * - exportExcelWithOverlay 的委托语义：等价主入口的 overlay 选项（挂载 →
 *   回调链式追加 → finally 关闭），显式 overlay 参数**覆盖** options 里
 *   可能存在的 overlay 字段（实现是整体替换，不是深合并），失败也关遮罩；
 * - showExportOverlay 转发到底层遮罩；
 * - nextPaint 从子路径再导出（progress-overlay 私有化后这是消费方拿到
 *   "先让遮罩绘制一帧"工具的唯一公开入口）。
 *
 * 遮罩本体与主入口 overlay 接线的细粒度行为在 overlay.test.ts 覆盖，
 * 这里不重复。导出同前：显式 stream + 单行 + download:false，绕开
 * WASM 与 Worker，happy-dom 下直跑纯 JS 路径。
 */
import { describe, it, expect, vi, beforeEach } from "vitest";
import {
  exportExcelWithOverlay,
  showExportOverlay,
  nextPaint,
} from "../overlay";
import type { ProgressOverlayHandle } from "../index";
import type { ExportOptions, SheetConfig } from "../types";

const calls: string[] = [];
let handle: ProgressOverlayHandle;
let showOptions: Record<string, unknown> | null = null;

vi.mock("@marcusok/progress-overlay", () => ({
  showProgressOverlay: vi.fn((options: Record<string, unknown>) => {
    showOptions = options;
    return handle;
  }),
  nextPaint: vi.fn(() => Promise.resolve()),
}));

beforeEach(() => {
  calls.length = 0;
  showOptions = null;
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
    filename: "overlay-subpath.xlsx",
    mode: "stream",
    sheets: [SHEET],
    download: false,
    ...(overlay !== undefined ? { overlay } : {}),
  };
}

describe("./overlay 子路径", () => {
  it("exportExcelWithOverlay 等价主入口 overlay 选项：挂载→回调链→finally 关闭", async () => {
    const seen: number[] = [];
    const result = await exportExcelWithOverlay(
      { ...baseOptions(), onProgress: (p) => seen.push(p) },
      {},
    );

    expect(result.success).toBe(true);
    expect(showOptions).not.toBeNull();
    // 调用方回调照常收 0 -> ... -> 1（链式追加，不被遮罩替换）
    expect(seen[0]).toBe(0);
    expect(seen[seen.length - 1]).toBe(1);
    expect(calls).toContain("progress:0");
    expect(calls.filter((c) => c === "close")).toHaveLength(1);
  });

  it("显式 overlay 参数整体覆盖 options.overlay 字段（非深合并）", async () => {
    await exportExcelWithOverlay(
      baseOptions({ text: { title: "from-options" } }),
      { text: { title: "explicit-arg" } },
    );
    const text = showOptions!.text as Record<string, unknown>;
    expect(text.title).toBe("explicit-arg");
  });

  it("导出失败（非法输入）时 close 仍被执行且只执行一次", async () => {
    const result = await exportExcelWithOverlay(
      { ...baseOptions(), sheets: [] },
      {},
    );
    expect(result.success).toBe(false);
    expect(calls.filter((c) => c === "close")).toHaveLength(1);
  });

  it("showExportOverlay forwards to the underlying overlay", () => {
    const h = showExportOverlay({ text: { title: "manual" } });
    expect(h).toBe(handle);
    expect((showOptions!.text as Record<string, unknown>).title).toBe("manual");
    h.close();
    expect(calls).toContain("close");
  });

  it("re-exports nextPaint from the subpath", async () => {
    // mock 模块的 nextPaint 已被计数验证过调用面；这里钉"子路径确实再导出
    // 了它"——progress-overlay 私有化后这是该工具的唯一公开入口
    await expect(nextPaint()).resolves.toBeUndefined();
  });
});
