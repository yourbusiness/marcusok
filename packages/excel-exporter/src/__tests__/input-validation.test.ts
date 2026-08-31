import { describe, it, expect, vi } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { exportAsStream } from "../streaming-builder";
import { exportExcel } from "../index";
import type { SheetConfig } from "../types";
// Node's fetch rejects file://, so sync-init the WASM (see setup.ts);
// importing the module already runs the side-effectful bootstrap.
import { readBuffer } from "./setup";

// Invalid user input must fail with a specific error on every path, instead
// of silently producing a workbook Excel flags as corrupt (the pre-fix
// behavior for NaN/Infinity values, zero-span or out-of-bounds merges, and
// duplicate sheet names).
const baseSheet = (over: Partial<SheetConfig> = {}): SheetConfig => ({
  name: "S",
  columns: [
    { key: "a", header: "A" },
    { key: "b", header: "B" },
  ],
  data: [
    { a: 1, b: 2 },
    { a: 3, b: 4 },
  ],
  ...over,
});

describe("non-finite numbers (NaN/Infinity)", () => {
  it("stream path writes visible strings, not illegal <v>NaN</v> XML", async () => {
    const { bytes } = await exportAsStream([
      baseSheet({
        data: [{ a: NaN, b: Infinity }],
      }),
    ]);
    // Inspect the raw XML: xsd:double only allows finite values, so
    // <v>NaN</v> / <v>Infinity</v> would corrupt the workbook.
    const sheetXml = strFromU8(unzipSync(bytes)["xl/worksheets/sheet1.xml"]);
    expect(sheetXml).not.toContain("<v>NaN</v>");
    expect(sheetXml).not.toContain("<v>Infinity</v>");
    expect(sheetXml).not.toContain('t="n"');

    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("S")!;
    expect(ws.cell("A2").value).toBe("NaN");
    expect(ws.cell("B2").value).toBe("Infinity");
  });

  it("workbook path (main mode) writes the same visible strings", async () => {
    const r = await exportExcel({
      filename: "nan-main",
      download: false,
      mode: "main",
      sheets: [baseSheet({ data: [{ a: NaN, b: Infinity }] })],
    });
    expect(r.success).toBe(true);
    expect(r.engine).toBe("modern-xlsx");
    // Same raw-XML guard as the stream case: no illegal <v>NaN</v> number cell.
    const bytes = new Uint8Array(await r.blob!.arrayBuffer());
    const sheetXml = strFromU8(unzipSync(bytes)["xl/worksheets/sheet1.xml"]);
    expect(sheetXml).not.toContain("<v>NaN</v>");
    expect(sheetXml).not.toContain("<v>Infinity</v>");
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("S")!;
    expect(ws.cell("A2").value).toBe("NaN");
    expect(ws.cell("B2").value).toBe("Infinity");
  });

  it("SheetJS fallback writes the same visible strings", async () => {
    vi.stubGlobal("WebAssembly", undefined);
    try {
      const r = await exportExcel({
        filename: "nan-fallback",
        download: false,
        sheets: [baseSheet({ data: [{ a: NaN, b: Infinity }] })],
      });
      expect(r.engine).toBe("sheetjs");
      expect(r.success).toBe(true);
      const wb = await readBuffer(new Uint8Array(await r.blob!.arrayBuffer()));
      const ws = wb.getSheet("S")!;
      expect(ws.cell("A2").value).toBe("NaN");
      expect(ws.cell("B2").value).toBe("Infinity");
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

describe("merge range validation", () => {
  it.each([
    {
      bad: { row: 0, col: 0, rowspan: 0, colspan: 1 },
      msg: /rowspan\/colspan must be >= 1/,
    },
    {
      bad: { row: 0, col: 1, rowspan: 1, colspan: 2 },
      msg: /exceeds the 2 leaf columns/,
    },
    {
      bad: { row: 1, col: 0, rowspan: 2, colspan: 1 },
      msg: /exceeds the 2 data rows/,
    },
    {
      bad: { row: 0, col: -1, rowspan: 1, colspan: 1 },
      msg: /row\/col must be >= 0/,
    },
  ])("stream path rejects $bad", async ({ bad, msg }) => {
    await expect(
      exportAsStream([baseSheet({ merges: [bad] })]),
    ).rejects.toThrow(msg);
  });

  it("stream path rejects overlapping ranges", async () => {
    await expect(
      exportAsStream([
        baseSheet({
          merges: [
            { row: 0, col: 0, rowspan: 2, colspan: 1 },
            { row: 1, col: 0, rowspan: 1, colspan: 1 },
          ],
        }),
      ]),
    ).rejects.toThrow(/overlaps merge #0/);
  });

  it("workbook path (via exportExcel) reports the same failure", async () => {
    const r = await exportExcel({
      filename: "merge-main",
      download: false,
      mode: "main",
      sheets: [
        baseSheet({ merges: [{ row: 0, col: 0, rowspan: 0, colspan: 1 }] }),
      ],
    });
    // The pre-flight check in exportExcel fails the call directly (previously:
    // workbook build threw -> SheetJS fallback re-validated -> failed too,
    // after one wasted fallback attempt and a misleading warn).
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/rowspan\/colspan must be >= 1/);
  });

  it("invalid input fails immediately without the SheetJS fallback", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "preflight-no-fallback",
        download: false,
        mode: "main",
        sheets: [
          baseSheet({
            merges: [{ row: 0, col: 0, rowspan: 0, colspan: 1 }],
          }),
        ],
      });
      expect(r.success).toBe(false);
      // No engine ran: the failure came from the entry-point pre-flight, not
      // from a degraded (style-less) SheetJS attempt.
      expect(r.engine).toBeUndefined();
      expect(r.error?.message).toMatch(/rowspan\/colspan must be >= 1/);
      // And no "Falling back to SheetJS" warn may be printed for input errors.
      expect(warn).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it("SheetJS fallback reports the same failure", async () => {
    vi.stubGlobal("WebAssembly", undefined);
    try {
      const r = await exportExcel({
        filename: "merge-fallback",
        download: false,
        sheets: [
          baseSheet({
            merges: [{ row: 0, col: 1, rowspan: 1, colspan: 5 }],
          }),
        ],
      });
      expect(r.success).toBe(false);
      expect(r.error?.message).toMatch(/exceeds the 2 leaf columns/);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("still accepts valid merges (regression)", async () => {
    const { bytes } = await exportAsStream([
      baseSheet({ merges: [{ row: 0, col: 0, rowspan: 2, colspan: 1 }] }),
    ]);
    const wb = await readBuffer(bytes);
    expect(wb.getSheet("S")!.mergeCells).toContain("A2:A3");
  });
});

describe("duplicate sheet names", () => {
  it("stream path rejects duplicates before writing any XML", async () => {
    await expect(
      exportAsStream([baseSheet({ name: "S" }), baseSheet({ name: "S" })]),
    ).rejects.toThrow(/duplicate sheet name "S"/);
  });

  it("workbook path degrades to SheetJS which reports the same failure", async () => {
    const r = await exportExcel({
      filename: "dup-main",
      download: false,
      mode: "main",
      sheets: [baseSheet({ name: "S" }), baseSheet({ name: "S" })],
    });
    // Previously: modern-xlsx threw "already exists" -> fallback let SheetJS
    // silently rename to S_1. Now the pre-flight check (and, for direct
    // builders, every path) fails with a clear duplicate error.
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/duplicate sheet name "S"/);
  });

  it("SheetJS fallback rejects duplicates too", async () => {
    vi.stubGlobal("WebAssembly", undefined);
    try {
      const r = await exportExcel({
        filename: "dup-fallback",
        download: false,
        sheets: [baseSheet({ name: "S" }), baseSheet({ name: "S" })],
      });
      expect(r.success).toBe(false);
      expect(r.error?.message).toMatch(/duplicate sheet name "S"/);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

describe("stream feature warnings on nested columns", () => {
  it("warns about width/style/headerStyle set on deep child nodes", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      await exportAsStream([
        baseSheet({
          columns: [
            { key: "a", header: "A" },
            {
              header: "Group",
              headerStyle: { font: { bold: true } },
              children: [
                {
                  key: "b",
                  header: "B",
                  width: 20,
                  style: { font: { italic: true } },
                },
              ],
            },
          ],
        }),
      ]);
      // Pre-fix, the top-level-only scan saw none of these and the features
      // were dropped silently, breaking the documented "dropped with a
      // warning" contract (README stream-mode notes).
      const messages = warn.mock.calls.map((c) => String(c[0])).join("\n");
      expect(messages).toContain("width");
      expect(messages).toContain("style");
      expect(messages).toContain("headerStyle");
    } finally {
      warn.mockRestore();
    }
  });

  it("does not warn when no features are configured", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      await exportAsStream([baseSheet()]);
      expect(warn).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });
});

describe("empty columns", () => {
  it("exportExcel fails fast with a clear error (no SheetJS fallback)", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "empty-columns",
        download: false,
        mode: "main",
        // Pre-fix this crashed the Workbook autoFilter layout with a cryptic
        // TypeError (encodeCellRef(0, -1) -> "@1"), then silently degraded to
        // SheetJS. The pre-flight check must fail it directly instead.
        sheets: [baseSheet({ columns: [], autoFilter: true })],
      });
      expect(r.success).toBe(false);
      expect(r.engine).toBeUndefined();
      expect(r.error?.message).toMatch(/at least one column/);
      expect(warn).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it("stream path rejects empty columns with the same error", async () => {
    await expect(exportAsStream([baseSheet({ columns: [] })])).rejects.toThrow(
      /at least one column/,
    );
  });
});
