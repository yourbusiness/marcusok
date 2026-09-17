import { describe, it, expect, vi } from "vitest";
import { exportExcel } from "../index";
import type { ExportOptions, SheetConfig } from "../types";
import { readBuffer } from "./setup";

// The terminal degradation runs the pure-JS fast stream on the main thread
// (WASM unsupported / every WASM-dependent route failed). These tests stub
// WebAssembly away and drive it through the public exportExcel entry, so the
// routing, the degradation markers and the progress contract are exercised
// end to end.

function fallbackOptions(
  filename: string,
  sheets: SheetConfig[],
  onProgress?: (p: number) => void,
): ExportOptions {
  return { filename, download: false, sheets, onProgress };
}

/** exportExcel with WebAssembly unsupported: the stream fallback must run. */
async function exportWithFallback(
  filename: string,
  sheets: SheetConfig[],
  onProgress?: (p: number) => void,
) {
  vi.stubGlobal("WebAssembly", undefined);
  try {
    return await exportExcel(fallbackOptions(filename, sheets, onProgress));
  } finally {
    vi.unstubAllGlobals();
  }
}

const basicSheet: SheetConfig = {
  name: "Sheet1",
  columns: [
    { prop: "name", label: "Name" },
    { prop: "value", label: "Value" },
  ],
  data: [
    { name: "Alice", value: 10 },
    { name: "Bob", value: 20 },
  ],
};

describe("stream fallback (terminal degradation, WebAssembly unavailable)", () => {
  it("exports a real xlsx blob on the stream engine with degradation markers", async () => {
    const r = await exportWithFallback("fallback-direct", [basicSheet]);

    expect(r.success).toBe(true);
    expect(r.engine).toBe("modern-xlsx");
    expect(r.mode).toBe("stream");
    expect(r.rowCount).toBe(2);
    expect(r.blob).toBeInstanceOf(Blob);
    expect(r.blob!.size).toBeGreaterThan(0);
    // Styles are stripped in the fallback; the soft error signals that and
    // carries the degradation reason through programmatically.
    expect(r.error).toBeInstanceOf(Error);
    expect(r.error!.message).toMatch(/styles stripped/i);
    expect(r.error!.message).toContain("Reason: WebAssembly not supported");
  });

  it("emits the documented 0 -> 1 onProgress pair", async () => {
    const progress: number[] = [];
    const r = await exportWithFallback("fallback-progress", [basicSheet], (p) =>
      progress.push(p),
    );
    expect(r.success).toBe(true);
    expect(progress).toEqual([0, 1]);
  });

  it("writes grouped headers and data merges in the fallback", async () => {
    const grouped: SheetConfig = {
      name: "Sheet1",
      columns: [
        { prop: "product", label: "产品" },
        {
          label: "收入情况",
          children: [
            {
              label: "本月",
              children: [
                { prop: "m_qty", label: "数量" },
                { prop: "m_amt", label: "金额" },
              ],
            },
          ],
        },
      ],
      data: [{ product: "A", m_qty: 1, m_amt: 2 }],
      merges: [{ row: 0, col: 0, rowspan: 1, colspan: 2 }],
    };

    const r = await exportWithFallback("fallback-grouped", [grouped]);
    expect(r.success).toBe(true);

    const wb = await readBuffer(new Uint8Array(await r.blob!.arrayBuffer()));
    const ws = wb.getSheet("Sheet1")!;
    // 3 header rows (收入情况 > 本月 > 数量/金额) + 1 data row.
    expect(ws.rowCount).toBe(4);
    expect(ws.cell("A1").value).toBe("产品");
    expect(ws.cell("B1").value).toBe("收入情况");
    expect(ws.cell("B2").value).toBe("本月");
    expect(ws.cell("B3").value).toBe("数量");
    // Leaf header A1:A3, groups B1:C1 / B2:C2, data merge A4:B4 (row 0 data-relative).
    for (const range of ["A1:A3", "B1:C1", "B2:C2", "A4:B4"]) {
      expect(ws.mergeCells).toContain(range);
    }
  });

  it("honours multiple sheets in rowCount accounting", async () => {
    const r = await exportWithFallback("fallback-multi", [
      basicSheet,
      {
        name: "Sheet2",
        columns: [{ prop: "x", label: "X" }],
        data: [{ x: 1 }, { x: 2 }, { x: 3 }],
      },
    ]);

    expect(r.success).toBe(true);
    expect(r.rowCount).toBe(5);
  });

  it("warns about the fallback and lists configured layout features as dropped", async () => {
    // Parity with the stream path's per-feature warnings: the fallback notice
    // must say styles are stripped, and the stream build beneath it reports
    // width/freezeRows/autoFilter as unsupported instead of dropping them
    // silently.
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportWithFallback("fallback-dropped-features", [
        {
          ...basicSheet,
          freezeRows: 1,
          autoFilter: true,
          columns: basicSheet.columns.map((c) => ({ ...c, width: 12 })),
        },
      ]);
      expect(r.success).toBe(true);

      const messages = warn.mock.calls.map((c) => String(c[0])).join("\n");
      expect(messages).toContain("style-less fast stream");
      expect(messages).toContain("Reason: WebAssembly not supported");
      expect(messages).toContain("freezeRows");
      expect(messages).toContain("autoFilter");
      expect(messages).toContain("width");
    } finally {
      warn.mockRestore();
    }
  });

  it("keeps the plain fallback warning when no layout features are configured", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportWithFallback("fallback-no-features", [basicSheet]);
      expect(r.success).toBe(true);
      const messages = warn.mock.calls.map((c) => String(c[0])).join("\n");
      expect(messages).toContain("style-less fast stream");
      expect(messages).not.toContain("features not supported");
    } finally {
      warn.mockRestore();
    }
  });
});

// Node stream 路由：首次尝试就是 fast stream，build 期错误即终局——
// exportExcel 的 catch 必须直接失败，而不是经 finishWithStream 对同一
// 确定性失败再跑一遍（与 worker 链 retryOnMainThread 的防护对齐，见
// index.ts）。
describe("Node stream route failure is terminal (no doomed re-run)", () => {
  it("fails once with the original error instead of re-running the stream", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const phases: string[] = [];
    const progress: number[] = [];
    try {
      // Node 的 stream 路由会执行函数式 format；让它抛错即可让（唯一的）
      // 构建尝试确定性失败。
      const boom: SheetConfig = {
        name: "S",
        columns: [
          {
            prop: "a",
            label: "A",
            format: () => {
              throw new Error("format exploded");
            },
          },
        ],
        data: [{ a: 1 }],
      };
      const r = await exportExcel({
        filename: "stream-terminal",
        download: false,
        mode: "stream",
        sheets: [boom],
        onPhase: (phase) => phases.push(phase),
        onProgress: (p) => progress.push(p),
      });

      expect(r.success).toBe(false);
      expect(r.error?.message).toBe("format exploded");
      // 构建尝试恰好一次：若白跑一遍 finishWithStream，会出现第二个 build
      // 阶段并打印兜底告警。
      expect(phases.filter((p) => p === "build")).toHaveLength(1);
      const messages = warn.mock.calls.map((c) => String(c[0])).join("\n");
      expect(messages).not.toContain("Falling back");
      // 进度契约：结尾的 1 恰好发一次。
      expect(progress.filter((p) => p === 1)).toHaveLength(1);
    } finally {
      warn.mockRestore();
    }
  });
});
