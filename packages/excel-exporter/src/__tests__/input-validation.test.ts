import { describe, it, expect, vi } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { exportAsStream } from "../streaming-builder";
import { WorkbookBuilder } from "../workbook-builder";
import { exportExcel } from "../index";
import { validateMerges, validateSheetLimits } from "../format-utils";
import type { ColumnConfig, SheetConfig } from "../types";
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
      expect(r.engine).toBe("fast-xlsx");
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

  it("validates a large disjoint merge list without quadratic blowup", () => {
    // 旧实现逐条与全部前序 merge 比较（40k 条约 1s，导出端到端 10s+）。
    // 扫描线版本应是毫秒级；阈值放宽到 500ms，只拦截复杂度回归。
    const n = 40_000;
    const merges = Array.from({ length: n }, (_, i) => ({
      row: i,
      col: 0,
      rowspan: 1,
      colspan: 2,
    }));
    const sheet: SheetConfig = {
      name: "S",
      columns: [
        { prop: "a", label: "A" },
        { prop: "b", label: "B" },
      ],
      data: new Array(n).fill({ a: 1, b: 2 }),
      merges,
    };
    const t0 = performance.now();
    expect(() => validateMerges(sheet, 2)).not.toThrow();
    expect(performance.now() - t0).toBeLessThan(500);
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

describe("null/undefined options and invalid mode", () => {
  // 入口处在 try 之前解引用 options（发首个 0）：不守住这里，null options
  // 会以裸 TypeError reject，绕过结构化 { success: false } 契约。
  it("null/undefined options resolves with a structured failure (no throw)", async () => {
    for (const bad of [null, undefined]) {
      const r = await exportExcel(bad as never);
      expect(r.success).toBe(false);
      expect(r.error?.message).toMatch(/options must be an object/);
    }
  });

  it("a misspelled mode is rejected instead of silently routing as auto", async () => {
    // pickMode 对未知值静默落入 auto 分支：用户的显式意图被无声忽略。
    const r = await exportExcel({
      filename: "bad-mode",
      download: false,
      // @ts-expect-error runtime JS callers can pass a misspelled mode
      mode: "Main",
      sheets: [baseSheet()],
    });
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(
      /mode must be one of "auto", "main", "worker", "stream" \(got "Main"\)/,
    );
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

  it("rejects case-variant duplicates like Excel does (workbook path)", async () => {
    // Excel 的表名唯一性大小写不敏感："S" 与 "s" 同名，放行会产出 Excel
    // 可能要求修复的工作簿而 success 仍为 true。
    const r = await exportExcel({
      filename: "dup-case-main",
      download: false,
      mode: "main",
      sheets: [baseSheet({ name: "S" }), baseSheet({ name: "s" })],
    });
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/duplicate sheet name "s"/);
  });

  it("rejects case-variant duplicates on the stream path too", async () => {
    await expect(
      exportAsStream([baseSheet({ name: "S" }), baseSheet({ name: "s" })]),
    ).rejects.toThrow(/duplicate sheet name "s"/);
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

  it("stream route does not warn (width) for the injected index column's default width", async () => {
    // 回归：applyIndexColumn 曾给注入列无条件填默认 width 6，启用 indexColumn
    // 的 stream 导出即使从未配置任何 width 也必然触发 "features not
    // supported (width)" 误告警。注入列现在不带 width（默认宽由 Workbook
    // 路径消费时补），只有用户显式配置了宽度才告警。
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "idx-stream-default",
        download: false,
        mode: "stream",
        sheets: [baseSheet({ indexColumn: true })],
      });
      expect(r.success).toBe(true);
      expect(warn).not.toHaveBeenCalled();

      // 对照：用户显式给 indexColumn.width 仍要如实告警（它确实会被丢弃）
      const r2 = await exportExcel({
        filename: "idx-stream-explicit",
        download: false,
        mode: "stream",
        sheets: [baseSheet({ indexColumn: { width: 8 } })],
      });
      expect(r2.success).toBe(true);
      const messages = warn.mock.calls.map((c) => String(c[0])).join("\n");
      expect(messages).toContain("width");
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
    {
      label: "textRotation above the documented 0-180 range",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            style: { alignment: { textRotation: 270 } },
          },
        ],
      },
      message:
        /column "A" style: alignment\.textRotation must be an integer between 0 and 180/,
    },
    {
      label: "textRotation fractional",
      sheet: {
        headerStyle: { alignment: { textRotation: 45.5 } },
        columns: [{ prop: "a", label: "A" }],
      },
      message:
        /sheet "S" headerStyle: alignment\.textRotation must be an integer between 0 and 180/,
    },
    {
      label: "font.size non-positive",
      sheet: {
        dataStyle: { font: { size: 0 } },
        columns: [{ prop: "a", label: "A" }],
      },
      message:
        /sheet "S" dataStyle: font\.size must be a finite positive number/,
    },
    {
      // 直通引擎会照单写进 styles.xml（<color rgb="notacolor"/>），产出
      // Excel 可能要求修复的工作簿而 success 仍 true——与表名/merge 同口径
      // 前置拦截（OOXML 的 rgb 是 ST_UnsignedIntHex）。
      label: "font.color not 6-digit hex",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            style: { font: { color: "notacolor" as unknown as string } },
          },
        ],
      },
      message: /column "A" style: font\.color must be a 6-digit RGB hex string/,
    },
    {
      label: "fill.fgColor 8-digit ARGB rejected (documented shape is RGB)",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            style: {
              fill: {
                pattern: "solid",
                fgColor: "FFFF0000" as unknown as string,
              },
            },
          },
        ],
      },
      message:
        /column "A" style: fill\.fgColor must be a 6-digit RGB hex string/,
    },
    {
      label: "border style not an OOXML enum value",
      sheet: {
        headerStyle: {
          border: { bottom: { style: "wavy" as unknown as "thin" } },
        },
        columns: [{ prop: "a", label: "A" }],
      },
      message:
        /sheet "S" headerStyle: border\.bottom\.style must be one of the OOXML border styles/,
    },
    {
      label: "border side color validated too",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            style: {
              border: {
                top: { style: "thin", color: "red" as unknown as string },
              },
            },
          },
        ],
      },
      message:
        /column "A" style: border\.top\.color must be a 6-digit RGB hex string/,
    },
    {
      label: "font.size NaN (JS caller)",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            headerStyle: { font: { size: Number.NaN } },
          },
        ],
      },
      message:
        /column "A" headerStyle: font\.size must be a finite positive number/,
    },
    {
      // 拼错的判别值原先落进 applyFormat 的 default 分支被 toStr 字符串化：
      // 数字列整列变文本、success:true 零告警。
      label: "format.type misspelled (JS caller)",
      sheet: {
        columns: [
          { prop: "a", label: "A", format: { type: "percent" } as never },
        ],
      },
      message: /column "A" format\.type must be one of/,
    },
    {
      label: "format.type undefined (JS caller)",
      sheet: {
        columns: [{ prop: "a", label: "A", format: {} as never }],
      },
      message: /column "A" format\.type must be one of/,
    },
    {
      label: "enum map missing (JS caller)",
      sheet: {
        columns: [{ prop: "a", label: "A", format: { type: "enum" } as never }],
      },
      message: /column "A" format\.map must be an object/,
    },
    {
      label: "enum map null (JS caller)",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            format: { type: "enum", map: null } as never,
          },
        ],
      },
      message: /column "A" format\.map must be an object/,
    },
    {
      label: "date pattern non-string (JS caller)",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            format: { type: "date", pattern: 123 } as never,
          },
        ],
      },
      message: /column "A" format\.pattern must be a string/,
    },
    {
      label: "padding fill non-string (JS caller)",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            format: { type: "padding", fill: 0, length: 4 } as never,
          },
        ],
      },
      message: /column "A" format\.fill must be a string/,
    },
    {
      label: "padding fill missing (JS caller)",
      sheet: {
        columns: [
          {
            prop: "a",
            label: "A",
            format: { type: "padding", length: 4 } as never,
          },
        ],
      },
      message: /column "A" format\.fill must be a string/,
    },
    {
      // 非数组真值原先绕过 ?.length 放行，随后在两路消费端以裸 TypeError
      // 失败，且降级链把 stream 全量构建白跑一遍。
      label: "merges non-array object (JS caller)",
      sheet: { merges: {} as never },
      message: /sheet "S" merges must be an array of MergeRange/,
    },
    {
      // 数组元素为 null：旧实现直接解引用 c.children，抛裸 TypeError，
      // 没有表名/下标定位。
      label: "null column entry (JS caller)",
      sheet: { columns: [null] as never },
      message: /sheet "S" columns\[0\] must be an object/,
    },
    {
      label: "null child column entry (JS caller)",
      sheet: {
        columns: [{ prop: "g", label: "G", children: [null] }] as never,
      },
      message: /sheet "S" columns\[0\]\.children\[0\] must be an object/,
    },
    {
      label: "null merge entry (JS caller)",
      sheet: { merges: [null] as never },
      message: /sheet "S" merge #0 must be an object/,
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

  it("style numeric fields at their documented boundaries stay legal", async () => {
    // textRotation 0/180（types.ts 注释承诺的区间端点）与正的 font.size
    // 都是合法输入，前置校验不得误伤。
    const r = await exportExcel({
      filename: "style-validation-legal",
      download: false,
      mode: "main",
      sheets: [
        baseSheet({
          headerStyle: { alignment: { textRotation: 180 } },
          columns: [{ prop: "a", label: "A", style: { font: { size: 12 } } }],
        }),
      ],
    });
    expect(r.success).toBe(true);
    expect(r.error).toBeUndefined();
  });

  it("every documented format.type variant stays legal (no false positive)", async () => {
    // 枚举校验按 FormatSpec 的全部五个变体收窄，任一合法 spec 被误伤都会
    // 在这里红：date/datetime 的 pattern 可省略（有默认），enum 的 map 是
    // 必填对象，padding 的 fill/length 必填。
    const r = await exportExcel({
      filename: "format-type-legal",
      download: false,
      mode: "main",
      sheets: [
        baseSheet({
          columns: [
            {
              prop: "a",
              label: "A",
              format: { type: "datetime", pattern: "yyyy-MM-dd HH:mm" },
            },
            {
              prop: "b",
              label: "B",
              format: { type: "enum", map: { x: "是" }, fallback: "否" },
            },
          ],
        }),
      ],
    });
    expect(r.success).toBe(true);
    expect(r.error).toBeUndefined();
  });
});

describe("indexColumn validation", () => {
  it("rejects a non-integer or negative start with a structured error", async () => {
    for (const start of [1.5, -1]) {
      const r = await exportExcel({
        filename: "bad-start",
        download: false,
        sheets: [
          baseSheet({
            indexColumn: { start },
          }),
        ],
      });
      expect(r.success).toBe(false);
      expect(r.error?.message).toMatch(
        /indexColumn\.start must be a non-negative integer/,
      );
    }
  });

  it("rejects a user column claiming the reserved index prop (structured, not thrown)", async () => {
    const r = await exportExcel({
      filename: "prop-clash",
      download: false,
      sheets: [
        baseSheet({
          indexColumn: true,
          columns: [
            { prop: "__index__", label: "我的序号" },
            { prop: "a", label: "A" },
          ],
        }),
      ],
    });
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/is reserved for the index column/);
  });

  // 回归：构建器对 INDEX_PROP 的特判是**无条件**的（只按 prop 判定即写行号、
  // 永不读 data），而冲突检查原先只在 applyIndexColumn 内——它在未启用
  // indexColumn 时直接 return sheet。于是不开 indexColumn 的用户自带的
  // `__index__` 列会被行号静默顶替、数据丢失且 success 仍为 true。检查上移到
  // validateInput 后，两条路由都必须在构建前以结构化错误失败。
  it("rejects the reserved index prop even when indexColumn is not enabled", async () => {
    for (const mode of ["main", "stream"] as const) {
      const r = await exportExcel({
        filename: "prop-clash-no-index",
        download: false,
        mode,
        sheets: [
          baseSheet({
            columns: [
              { prop: "__index__", label: "ID" },
              { prop: "a", label: "A" },
            ],
          }),
        ],
      });
      expect(r.success).toBe(false);
      expect(r.error?.message).toMatch(/is reserved for the index column/);
    }
  });

  // 回归：序号列在 validateInput 之后才由 applyIndexColumn 注入，width 曾因此
  // 绕过早前的 col.width 校验——NaN 会以 serde 错误让整份导出降级为无样式
  // stream（success 仍为 true），与 >=50k 路由静默忽略 width 的行为分裂。
  it("rejects an invalid width instead of degrading the export", async () => {
    for (const width of [Number.NaN, -5, Number.POSITIVE_INFINITY]) {
      const r = await exportExcel({
        filename: "bad-index-width",
        download: false,
        mode: "main",
        sheets: [baseSheet({ indexColumn: { width } })],
      });
      expect(r.success).toBe(false);
      expect(r.error?.message).toMatch(
        /indexColumn\.width must be a finite non-negative number/,
      );
    }
  });

  it("accepts width 0 (hidden column) on the index column", async () => {
    const r = await exportExcel({
      filename: "zero-index-width",
      download: false,
      mode: "main",
      sheets: [baseSheet({ indexColumn: { width: 0 } })],
    });
    expect(r.success).toBe(true);
    expect(r.error).toBeUndefined();
  });

  it("accepts start 0 and a valid indexColumn on the workbook path", async () => {
    const r = await exportExcel({
      filename: "ok-start",
      download: false,
      mode: "main",
      sheets: [
        baseSheet({
          indexColumn: { start: 0 },
        }),
      ],
    });
    expect(r.success).toBe(true);
    expect(r.error).toBeUndefined();
  });
});

describe("indexColumn shape validation", () => {
  it("rejects non-boolean non-object values (null / string / number) with a structured error", async () => {
    for (const indexColumn of [null, "yes", 1]) {
      const r = await exportExcel({
        filename: "bad-index-column",
        download: false,
        sheets: [
          baseSheet({
            indexColumn: indexColumn as unknown as true,
          }),
        ],
      });
      expect(r.success).toBe(false);
      expect(r.error?.message).toMatch(
        /indexColumn must be a boolean or an options object/,
      );
    }
  });
});

// Excel 的硬上限（ECMA-376）：每表 1,048,576 行（表头占行）与 16,384 列。
// 超限输入此前在 stream 路径照常写出全部 <row>、zip 成功，Excel 打开即报
// 修复而 success 仍为 true——正是 validateInput 设计哲学要消灭的失败类型。
describe("Excel sheet dimension limits", () => {
  it("rejects data that overflows the 1,048,576-row limit (header included)", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const r = await exportExcel({
        filename: "too-many-rows",
        download: false,
        mode: "main",
        sheets: [
          baseSheet({
            // 共享引用填充，只为过长度校验；校验通过前不会触碰元素
            data: new Array(1_048_576).fill({ a: 1, b: 2 }),
          }),
        ],
      });
      expect(r.success).toBe(false);
      expect(r.engine).toBeUndefined();
      expect(r.error?.message).toMatch(/1,048,576-row limit/);
      expect(warn).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it("rejects more than 16,384 effective columns, counting the injected index column", async () => {
    const manyColumns = Array.from({ length: 16_384 }, (_, i) => ({
      prop: `c${i}`,
      label: `C${i}`,
    }));
    const r = await exportExcel({
      filename: "too-many-cols",
      download: false,
      mode: "main",
      // 16,384 用户列 + 序号列 = 16,385 有效列，越界
      sheets: [baseSheet({ columns: manyColumns, indexColumn: true })],
    });
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/16,384-column limit/);
  });

  it("stream path rejects the same overflow", async () => {
    await expect(
      exportAsStream([
        baseSheet({ data: new Array(1_048_576).fill({ a: 1, b: 2 }) }),
      ]),
    ).rejects.toThrow(/1,048,576-row limit/);
    await expect(
      exportAsStream([
        baseSheet({
          columns: Array.from({ length: 16_385 }, (_, i) => ({
            prop: `c${i}`,
            label: `C${i}`,
          })),
        }),
      ]),
    ).rejects.toThrow(/16,384-column limit/);
  });

  it("boundary values stay legal", () => {
    const mk = (rows: number): SheetConfig => ({
      name: "S",
      columns: [{ prop: "a", label: "A" }],
      data: new Array(rows).fill({ a: 1 }),
    });
    // 行：表头 1 行 + 数据 1,048,575 行 = 恰好 1,048,576，合法；多一行越界
    expect(() => validateSheetLimits(mk(1_048_575), 1, 1)).not.toThrow();
    expect(() => validateSheetLimits(mk(1_048_576), 1, 1)).toThrow(
      /1,048,576-row limit/,
    );
    // 列：16,384 合法，16,385 越界
    expect(() => validateSheetLimits(mk(1), 16_384, 1)).not.toThrow();
    expect(() => validateSheetLimits(mk(1), 16_385, 1)).toThrow(
      /16,384-column limit/,
    );
  });
});

const C = (n: number): string => String.fromCharCode(n);

// XML 1.0 之外的控制字符/非字符/孤立代理项不能出现在 OOXML 中：modern-xlsx
// 对 sheet 名、font.name、numFormat、format.pattern 会原样写进 XML，导致
// Excel 打开即要求修复，而导出返回 success:true。这些字段无法静默清洗
// （改名/改格式码会改变业务语义），必须前置拒绝。
describe("XML-illegal characters in configuring strings", () => {
  const cases: {
    label: string;
    sheet: SheetConfig;
    message: RegExp;
  }[] = [
    {
      label: "sheet name U+0001",
      sheet: baseSheet({ name: `S${C(1)}X` }),
      message: /sheet name .* contains U\+0001/,
    },
    {
      label: "sheet name U+FFFE",
      sheet: baseSheet({ name: `S${C(0xfffe)}X` }),
      message: /sheet name .* contains U\+FFFE/,
    },
    {
      label: "sheet name lone surrogate",
      sheet: baseSheet({ name: `S${C(0xd800)}X` }),
      message: /sheet name .* contains U\+D800/,
    },
    {
      label: "font.name U+0001",
      sheet: baseSheet({
        columns: [
          { prop: "a", label: "A", style: { font: { name: `F${C(1)}X` } } },
        ],
      }),
      message: /font\.name contains U\+0001/,
    },
    {
      label: "numFormat U+FFFE",
      sheet: baseSheet({
        columns: [
          { prop: "a", label: "A", style: { numFormat: `0.00${C(0xfffe)}` } },
        ],
      }),
      message: /numFormat contains U\+FFFE/,
    },
    {
      label: "date pattern U+0001",
      sheet: baseSheet({
        columns: [
          {
            prop: "a",
            label: "A",
            format: { type: "date", pattern: `yyyy${C(1)}` },
          },
        ],
      }),
      message: /format\.pattern contains U\+0001/,
    },
  ];

  for (const { label, sheet, message } of cases) {
    it(`rejects ${label} before building (no engine call, no fallback)`, async () => {
      const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
      try {
        for (const mode of ["main", "stream"] as const) {
          const r = await exportExcel({
            filename: "xml-config",
            download: false,
            mode,
            sheets: [sheet],
          });
          expect(r.success).toBe(false);
          expect(r.error?.message).toMatch(message);
          expect(warn).not.toHaveBeenCalled();
        }
      } finally {
        warn.mockRestore();
      }
    });
  }
});

// 数据/表头文本是"可清洗"的：删除 XML 非法控制字符、孤立代理项替换为
// U+FFFD。两条导出路径必须写出相同内容，且不得出现原始非法字节。
describe("XML-safe text sanitization for data/headers", () => {
  const cases = [
    { label: "U+0001", text: `a${C(1)}b`, expected: "ab" },
    {
      label: "lone surrogate",
      text: `a${C(0xd800)}b`,
      expected: `a${C(0xfffd)}b`,
    },
  ];

  for (const { label, text, expected } of cases) {
    it(`sanitizes ${label} identically on workbook and stream paths`, async () => {
      for (const mode of ["main", "stream"] as const) {
        const r = await exportExcel({
          filename: "xml-data",
          download: false,
          mode,
          sheets: [
            {
              name: "S",
              columns: [{ prop: "v", label: `H${C(1)}X` }],
              data: [{ v: text }],
            },
          ],
        });
        expect(r.success).toBe(true);
        expect(r.error).toBeUndefined();
        const bytes = new Uint8Array(await r.blob!.arrayBuffer());
        const wb = await readBuffer(bytes);
        const ws = wb.getSheet("S")!;
        // 表头同样经过清洗；数据值按 XML 1.0 规则归一
        expect(ws.cell("A1").value).toBe("HX");
        expect(ws.cell("A2").value).toBe(expected);
        // 所有 XML 部件不得残留原始控制字符/非字符
        for (const [name, part] of Object.entries(unzipSync(bytes))) {
          if (!name.endsWith(".xml")) continue;
          expect(strFromU8(part)).not.toMatch(
            /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/,
          );
        }
      }
    });
  }
});

// TypeScript 只能约束 TS 调用方；JS 调用方的非法枚举/类型会被引擎原样
// 写进 styles.xml（Excel 判损），或在 serde 处失败并整份降级。这里统一
// 前置为结构化失败。
describe("style enum/type runtime validation (JS callers)", () => {
  const cases: {
    label: string;
    style: Record<string, unknown>;
    message: RegExp;
  }[] = [
    {
      label: "fill.pattern misspelled",
      style: { fill: { pattern: "bogus", fgColor: "FF0000" } },
      message: /fill\.pattern must be "solid" or "none"/,
    },
    {
      label: "horizontal alignment unsupported",
      style: { alignment: { horizontal: "middle" } },
      message: /alignment\.horizontal must be "left", "center" or "right"/,
    },
    {
      label: "vertical alignment unsupported",
      style: { alignment: { vertical: "middle" } },
      message: /alignment\.vertical must be "top", "center" or "bottom"/,
    },
    {
      label: "font.bold non-boolean",
      style: { font: { bold: "yes" } },
      message: /font\.bold must be a boolean/,
    },
    {
      label: "font.name non-string",
      style: { font: { name: 123 } },
      message: /font\.name must be a string/,
    },
    {
      label: "numFormat non-string",
      style: { numFormat: 123 },
      message: /numFormat must be a string/,
    },
    {
      label: "fill non-object",
      style: { fill: "solid" },
      message: /fill must be an object/,
    },
  ];

  for (const { label, style, message } of cases) {
    it(`rejects ${label} before building`, async () => {
      const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
      try {
        // 测试运行时校验：把非法样式以 Object.assign 注入（绕过类型系统，
        // 模拟 JS 调用方），避免为构造非法值写无意义的类型断言。
        const column: ColumnConfig = { prop: "a", label: "A" };
        Object.assign(column, { style });
        const r = await exportExcel({
          filename: "style-validation",
          download: false,
          mode: "main",
          sheets: [baseSheet({ columns: [column] })],
        });
        expect(r.success).toBe(false);
        expect(r.error?.message).toMatch(message);
        expect(warn).not.toHaveBeenCalled();
      } finally {
        warn.mockRestore();
      }
    });
  }

  it("keeps all documented enum values legal", async () => {
    const r = await exportExcel({
      filename: "style-validation-legal-2",
      download: false,
      mode: "main",
      sheets: [
        baseSheet({
          columns: [
            {
              prop: "a",
              label: "A",
              style: {
                font: { bold: true, italic: false, name: "Arial" },
                fill: { pattern: "none" },
                alignment: {
                  horizontal: "left",
                  vertical: "bottom",
                  wrapText: false,
                },
              },
            },
          ],
        }),
      ],
    });
    expect(r.success).toBe(true);
    expect(r.error).toBeUndefined();
  });

  it("low-level WorkbookBuilder rejects invalid styles too", async () => {
    const builder = await WorkbookBuilder.create();
    expect(() =>
      builder.addSheet(
        baseSheet({
          columns: [
            {
              prop: "a",
              label: "A",
              style: { fill: { pattern: "bogus" as never } },
            },
          ],
        }),
      ),
    ).toThrow(/fill\.pattern must be "solid" or "none"/);
  });
});
