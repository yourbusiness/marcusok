import { describe, expect, it, vi } from "vitest";
import { exportExcel } from "../index";
import { exportAsStream } from "../streaming-builder";

const sheets = [
  {
    name: "S",
    columns: [{ prop: "a", label: "A" }],
    data: [{ a: 1 }],
  },
];

describe("AbortSignal cancellation", () => {
  it("pre-aborted main export resolves a structured AbortError with 0 -> 1 progress", async () => {
    const ac = new AbortController();
    ac.abort();
    const progress: number[] = [];
    const r = await exportExcel({
      filename: "abort-main",
      download: false,
      mode: "main",
      signal: ac.signal,
      onProgress: (p) => progress.push(p),
      sheets,
    });
    expect(r.success).toBe(false);
    expect(r.error?.name).toBe("AbortError");
    // Cancellation happens before any engine work; no degradation ran.
    expect(r.engine).toBeUndefined();
    expect(progress).toEqual([0, 1]);
  });

  it("pre-aborted stream export resolves the same failure shape", async () => {
    const ac = new AbortController();
    ac.abort();
    const r = await exportExcel({
      filename: "abort-stream",
      download: false,
      mode: "stream",
      signal: ac.signal,
      sheets,
    });
    expect(r.success).toBe(false);
    expect(r.error?.name).toBe("AbortError");
    expect(r.engine).toBeUndefined();
  });

  it("exportAsStream honors an already-aborted signal", async () => {
    const ac = new AbortController();
    ac.abort();
    await expect(
      exportAsStream(sheets, undefined, ac.signal),
    ).rejects.toMatchObject({ name: "AbortError" });
  });

  it("aborting during the terminal stream fallback keeps the AbortError", async () => {
    // 无 WebAssembly → 直接降级到 fast stream。onProgress 中断开 signal：
    // runOnMainThread 返回结构化 AbortError 后，finishWithStream 不得把它
    // 覆盖成 "Fallback: styles stripped"（调用方靠 name 识别取消）。
    vi.stubGlobal("WebAssembly", undefined);
    try {
      const ac = new AbortController();
      const r = await exportExcel({
        filename: "abort-during-fallback",
        download: false,
        overlay: false,
        signal: ac.signal,
        onProgress: (p) => {
          if (p > 0 && p < 1) ac.abort();
        },
        sheets: [
          {
            name: "S",
            columns: [{ prop: "a", label: "A" }],
            data: Array.from({ length: 5000 }, (_, i) => ({ a: i })),
          },
        ],
      });
      expect(r.success).toBe(false);
      expect(r.error?.name).toBe("AbortError");
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
