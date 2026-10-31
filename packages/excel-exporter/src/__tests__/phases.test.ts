import { describe, expect, it, vi } from "vitest";
import { exportExcel } from "../index";
// Node's fetch rejects file://, so sync-init the WASM (see setup.ts).
import "./setup";

describe("exportExcel phase reporting (onPhase)", () => {
  it("reports init -> build in order on the main-thread path", async () => {
    const phases: string[] = [];
    const durations: number[] = [];
    const r = await exportExcel({
      filename: "phases-main",
      download: false,
      mode: "main",
      sheets: [
        {
          name: "S",
          columns: [{ prop: "x", label: "X" }],
          data: [{ x: 1 }, { x: 2 }],
        },
      ],
      onPhase: (phase, durationMs) => {
        phases.push(phase);
        durations.push(durationMs);
      },
    });
    expect(r.success).toBe(true);
    expect(r.engine).toBe("modern-xlsx");
    // download: false, so the download phase must not be reported.
    expect(phases).toEqual(["init", "build"]);
    for (const d of durations) {
      expect(Number.isFinite(d)).toBe(true);
      expect(d).toBeGreaterThanOrEqual(0);
    }
  });

  it("reports init -> build on the main-thread stream path", async () => {
    const phases: string[] = [];
    const r = await exportExcel({
      filename: "phases-stream",
      download: false,
      mode: "stream",
      sheets: [
        {
          name: "S",
          columns: [{ prop: "id", label: "ID" }],
          data: Array.from({ length: 100 }, (_, i) => ({ id: i })),
        },
      ],
      onPhase: (phase) => {
        phases.push(phase);
      },
    });
    expect(r.success).toBe(true);
    expect(r.mode).toBe("stream");
    expect(phases).toEqual(["init", "build"]);
  });

  it("does not report a download phase in Node even with download left on", async () => {
    // Node has no document, so no download happens and the phase must not be
    // reported (types.ts: "Not reported in Node (no document)"). Default
    // download is true -- do NOT pass download:false here.
    const phases: string[] = [];
    const r = await exportExcel({
      filename: "phases-node-download",
      mode: "main",
      sheets: [
        {
          name: "S",
          columns: [{ prop: "x", label: "X" }],
          data: [{ x: 1 }],
        },
      ],
      onPhase: (phase) => {
        phases.push(phase);
      },
    });
    expect(r.success).toBe(true);
    expect(phases).toEqual(["init", "build"]);
  });
});

// onProgress/onPhase 是调用方代码：其抛错（常见如回调驱动的 UI 组件已卸载）
// 不得影响导出本体——否则既突破"永远 resolve 结构化结果"的契约，又会把一次
// 已成功的导出误诊为构建失败而触发降级重建。抛错应仅 warn、导出照常完成。
describe("user callback error isolation", () => {
  const sheet = {
    name: "S",
    columns: [{ prop: "x", label: "X" }],
    data: [{ x: 1 }, { x: 2 }],
  };
  const boom = () => {
    throw new Error("ui component unmounted");
  };

  it("a throwing onProgress does not fail a main-thread export", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "cb-throw-main",
        download: false,
        mode: "main",
        sheets: [sheet],
        onProgress: boom,
      });
      expect(r.success).toBe(true);
      expect(r.blob).toBeDefined();
      expect(warn).toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it("a throwing onProgress does not fail the stream path", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "cb-throw-stream",
        download: false,
        mode: "stream",
        sheets: [sheet],
        onProgress: boom,
      });
      expect(r.success).toBe(true);
      expect(r.blob).toBeDefined();
    } finally {
      warn.mockRestore();
    }
  });

  it("an async onProgress rejection is contained (no unhandled rejection)", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "cb-async-reject",
        download: false,
        mode: "main",
        sheets: [sheet],
        // eslint-disable-next-line @typescript-eslint/no-misused-promises -- 有意模拟 async void 回调的 rejected promise
        onProgress: async () => {
          await Promise.resolve();
          throw new Error("async ui component unmounted");
        },
      });
      // 让 guardCallback 挂上的 rejection 处理器跑完（未接住会由 vitest
      // 报 unhandled rejection 直接失败）
      await new Promise((res) => setTimeout(res, 0));
      expect(r.success).toBe(true);
      expect(r.blob).toBeDefined();
      expect(warn).toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it("a throwing onPhase does not fail the export", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "cb-throw-phase",
        download: false,
        mode: "main",
        sheets: [sheet],
        onPhase: boom,
      });
      expect(r.success).toBe(true);
      expect(r.blob).toBeDefined();
    } finally {
      warn.mockRestore();
    }
  });

  it("failed exports still resolve a structured result when callbacks throw", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "",
        download: false,
        mode: "main",
        sheets: [sheet],
        onProgress: boom,
      });
      expect(r.success).toBe(false);
      expect(r.error?.message).toMatch(/filename must be a non-empty string/);
    } finally {
      warn.mockRestore();
    }
  });
});
