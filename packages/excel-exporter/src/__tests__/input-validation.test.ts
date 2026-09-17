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
    { prop: "a", label: "A" },
    { prop: "b", label: "B" },
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

  it("stream fallback writes the same visible strings", async () => {
    vi.stubGlobal("WebAssembly", undefined);
    try {
      const r = await exportExcel({
        filename: "nan-fallback",
        download: false,
        sheets: [baseSheet({ data: [{ a: NaN, b: Infinity }] })],
      });
      expect(r.engine).toBe("modern-xlsx");
      expect(r.mode).toBe("stream");
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
    // workbook build threw -> fallback re-validated -> failed too,
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
      // from a degraded (style-less) stream attempt.
      expect(r.engine).toBeUndefined();
      expect(r.error?.message).toMatch(/rowspan\/colspan must be >= 1/);
      // And no "Falling back" warn may be printed for input errors.
      expect(warn).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it("stream fallback reports the same failure", async () => {
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

describe("empty sheets array", () => {
  it("exportExcel fails fast with a clear error (no fallback, no corrupt file)", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "empty-sheets",
        download: false,
        mode: "main",
        // Pre-fix the Workbook build threw (modern-xlsx requires >= 1 sheet),
        // the export degraded to the fast stream and RESOLVED success: true
        // with a zero-sheet workbook Excel flags as corrupt.
        sheets: [],
      });
      expect(r.success).toBe(false);
      expect(r.blob).toBeUndefined();
      expect(r.error?.message).toMatch(/at least one sheet/);
      expect(warn).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it("stream path rejects an empty sheets array with the same error", async () => {
    await expect(exportAsStream([])).rejects.toThrow(/at least one sheet/);
  });
});

describe("missing or invalid filename", () => {
  // Pre-fix, filename was never validated: a JS caller omitting it got
  // success:true on the Node route, and on the browser route a masked
  // TypeError inside triggerDownload (caught as a cryptic warning) — no file
  // on disk either way. It now fails fast like every other structural input.
  it("resolves with a structured failure, keeping the 0 -> 1 progress contract", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const onProgress = vi.fn();
    try {
      const r = await exportExcel({
        // @ts-expect-error runtime JS callers can omit it
        filename: undefined,
        download: false,
        mode: "main",
        onProgress,
        sheets: [baseSheet()],
      });
      expect(r.success).toBe(false);
      expect(r.blob).toBeUndefined();
      expect(r.error?.message).toMatch(/filename must be a non-empty string/);
      expect(warn).not.toHaveBeenCalled();
      expect(onProgress.mock.calls.map((c) => c[0])).toEqual([0, 1]);
    } finally {
      warn.mockRestore();
    }
  });

  it("rejects an empty-string filename with the same error", async () => {
    const r = await exportExcel({
      filename: "",
      download: false,
      mode: "main",
      sheets: [baseSheet()],
    });
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/filename must be a non-empty string/);
  });
});

describe("structurally malformed sheets input", () => {
  // Pre-fix, `totalRows` was reduce()d before validateInput ran, so these
  // shapes rejected the promise with a raw TypeError instead of resolving
  // with the documented { success: false, error }.
  it("non-array `sheets` resolves with a structured failure (no throw)", async () => {
    const onProgress = vi.fn();
    const r = await exportExcel({
      filename: "no-sheets",
      download: false,
      onProgress,
      sheets: undefined as unknown as SheetConfig[],
    });
    expect(r.success).toBe(false);
    expect(r.blob).toBeUndefined();
    expect(r.error?.message).toMatch(/at least one sheet/);
    // The 0 -> 1 progress contract holds for failed exports too.
    expect(onProgress.mock.calls).toEqual([[0], [1]]);
  });

  it("a sheet without a columns array fails with a clear error", async () => {
    const r = await exportExcel({
      filename: "no-columns",
      download: false,
      mode: "main",
      sheets: [{ name: "S", data: [] } as unknown as SheetConfig],
    });
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/sheet "S" must have a columns array/);
  });

  it("a sheet without a data array fails with a clear error", async () => {
    const r = await exportExcel({
      filename: "no-data",
      download: false,
      mode: "main",
      sheets: [
        {
          name: "S",
          columns: [{ prop: "a", label: "A" }],
        } as unknown as SheetConfig,
      ],
    });
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/sheet "S" must have a data array/);
  });

  it("a null sheet entry fails with a clear error (no raw TypeError)", async () => {
    const r = await exportExcel({
      filename: "null-sheet",
      download: false,
      sheets: [null as unknown as SheetConfig],
    });
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/each sheet must be an object/);
  });
});

describe("duplicate sheet names", () => {
  it("stream path rejects duplicates before writing any XML", async () => {
    await expect(
      exportAsStream([baseSheet({ name: "S" }), baseSheet({ name: "S" })]),
    ).rejects.toThrow(/duplicate sheet name "S"/);
  });

  it("workbook path fails with the same duplicate error", async () => {
    const r = await exportExcel({
      filename: "dup-main",
      download: false,
      mode: "main",
      sheets: [baseSheet({ name: "S" }), baseSheet({ name: "S" })],
    });
    // Previously: modern-xlsx threw "already exists" -> the SheetJS fallback
    // silently renamed to S_1. Now the pre-flight check (and, for direct
    // builders, every path) fails with a clear duplicate error.
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/duplicate sheet name "S"/);
  });

  it("stream fallback rejects duplicates too", async () => {
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
            { prop: "a", label: "A" },
            {
              label: "Group",
              headerStyle: { font: { bold: true } },
              children: [
                {
                  prop: "b",
                  label: "B",
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
  it("exportExcel fails fast with a clear error (no fallback)", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "empty-columns",
        download: false,
        mode: "main",
        // Pre-fix this crashed the Workbook autoFilter layout with a cryptic
        // TypeError (encodeCellRef(0, -1) -> "@1"), then silently degraded to
        // the fallback. The pre-flight check must fail it directly instead.
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

// 数值型布局/格式字段的跨路径一致性校验：非法 width/freezeRows/decimals
// 此前在 Workbook 路径以晦涩的引擎 serde 错误失败并触发整份降级，在
// stream 路径却静默忽略或抛 RangeError——前置校验让两条路径同样直接失败。
describe("numeric field validation (width / freezeRows / format spec)", () => {
  const cases: Array<{
    label: string;
    sheet: Partial<SheetConfig>;
    message: RegExp;
  }> = [
    {
      label: "width NaN",
      sheet: { columns: [{ prop: "a", label: "A", width: NaN }] },
      message: /column "A" width must be a finite non-negative number/,
    },
    {
      label: "width negative",
      sheet: {
        columns: [
          { prop: "a", label: "A" },
          { prop: "b", label: "B", width: -5 },
        ],
      },
      message: /column "B" width must be a finite non-negative number/,
    },
    {
      label: "width non-number (JS caller)",
      sheet: {
        columns: [{ prop: "a", label: "A", width: "20" as unknown as number }],
      },
      message: /column "A" width must be a finite non-negative number/,
    },
    {
      label: "freezeRows fractional",
      sheet: { freezeRows: 1.5 },
      message: /sheet "S" freezeRows must be a non-negative integer/,
    },
    {
      label: "freezeRows negative",
      sheet: { freezeRows: -1 },
      message: /sheet "S" freezeRows must be a non-negative integer/,
    },
    {
      label: "decimals negative",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            format: { type: "number", decimals: -1 },
          },
        ],
      },
      message:
        /column "A" format\.decimals must be an integer between 0 and 100/,
    },
    {
      label: "decimals above toFixed limit",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            format: { type: "number", decimals: 105 },
          },
        ],
      },
      message:
        /column "A" format\.decimals must be an integer between 0 and 100/,
    },
    {
      label: "padding length negative",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            format: { type: "padding", fill: "0", length: -1 },
          },
        ],
      },
      message:
        /column "A" format\.length must be an integer between 0 and 10000/,
    },
  ];

  for (const { label, sheet, message } of cases) {
    it(`fails fast on ${label} (no engine call, no fallback)`, async () => {
      const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
      try {
        const r = await exportExcel({
          filename: "numeric-validation",
          download: false,
          mode: "main",
          sheets: [baseSheet(sheet)],
        });
        expect(r.success).toBe(false);
        expect(r.error?.message).toMatch(message);
        expect(warn).not.toHaveBeenCalled();
      } finally {
        warn.mockRestore();
      }
    });
  }

  it("width 0 stays legal (OOXML hides the column) and freezeRows 0 is a no-op", async () => {
    const r = await exportExcel({
      filename: "numeric-validation-legal",
      download: false,
      mode: "main",
      sheets: [
        baseSheet({
          freezeRows: 0,
          columns: [
            { prop: "a", label: "A", width: 0 },
            { prop: "b", label: "B" },
          ],
        }),
      ],
    });
    expect(r.success).toBe(true);
    expect(r.error).toBeUndefined();
  });
});
