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
    { key: "name", header: "Name" },
    { key: "value", header: "Value" },
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
        { key: "product", header: "产品" },
        {
          header: "收入情况",
          children: [
            {
              header: "本月",
              children: [
                { key: "m_qty", header: "数量" },
                { key: "m_amt", header: "金额" },
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
        columns: [{ key: "x", header: "X" }],
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
