import { describe, it, expect } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { exportAsStream } from "../streaming-builder";
import { readBuffer } from "./setup";

describe("exportAsStream round-trip", () => {
  it("produces a valid xlsx with correct row count and values", async () => {
    const sheets = [
      {
        name: "Data",
        columns: [
          { prop: "id", label: "ID" },
          { prop: "name", label: "Name" },
          { prop: "amount", label: "Amount" },
          { prop: "status", label: "Status" },
        ],
        data: Array.from({ length: 1000 }, (_, i) => ({
          id: i,
          name: `row_${i}`,
          amount: i * 1.5,
          status: i % 2 === 0 ? "paid" : "pending",
        })),
      },
    ];
    const { bytes, rowCount } = await exportAsStream(sheets);

    expect(bytes.slice(0, 4)).toEqual(new Uint8Array([0x50, 0x4b, 0x03, 0x04]));
    expect(rowCount).toBe(1000);

    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("Data")!;
    // header + 1000 data rows
    expect(ws.rowCount).toBe(1001);
    expect(ws.cell("A1").value).toBe("ID");
    expect(ws.cell("D1").value).toBe("Status");
    expect(String(ws.cell("A2").value)).toBe("0");
    expect(String(ws.cell("C1001").value)).toBe(String(999 * 1.5));
  });

  it("writes multi-row grouped headers and merges in stream mode", async () => {
    const { bytes } = await exportAsStream([
      {
        name: "Grouped",
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
              {
                label: "本年累计",
                children: [
                  { prop: "y_qty", label: "数量" },
                  { prop: "y_amt", label: "金额" },
                ],
              },
            ],
          },
        ],
        data: [{ product: "A", m_qty: 1, m_amt: 2, y_qty: 3, y_amt: 4 }],
      },
    ]);
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("Grouped")!;
    // 3 header rows + 1 data row.
    expect(ws.rowCount).toBe(4);
    expect(ws.cell("B1").value).toBe("收入情况");
    expect(ws.cell("B2").value).toBe("本月");
    expect(ws.cell("B3").value).toBe("数量");
    expect(ws.cell("E3").value).toBe("金额");
    expect(ws.cell("A4").value).toBe("A");
    for (const r of ["A1:A3", "B1:E1", "B2:C2", "D2:E2"]) {
      expect(ws.mergeCells).toContain(r);
    }
  });

  it("emits no mergeCells for a flat header", async () => {
    const { bytes } = await exportAsStream([
      {
        name: "S",
        columns: [
          { prop: "a", label: "A" },
          { prop: "b", label: "B" },
        ],
        data: [{ a: 1, b: 2 }],
      },
    ]);
    const wb = await readBuffer(bytes);
    // modern-xlsx returns null (not []) when the sheet has no <mergeCells>.
    expect(wb.getSheet("S")!.mergeCells ?? []).toHaveLength(0);
  });

  it("handles multi-sheet streaming", async () => {
    const { bytes, rowCount } = await exportAsStream([
      {
        name: "A",
        columns: [{ prop: "x", label: "X" }],
        data: [{ x: 1 }, { x: 2 }],
      },
      { name: "B", columns: [{ prop: "y", label: "Y" }], data: [{ y: 3 }] },
    ]);
    expect(rowCount).toBe(3);
    const wb = await readBuffer(bytes);
    expect(wb.sheetNames).toEqual(["A", "B"]);
    expect(wb.getSheet("A")!.rowCount).toBe(3);
    expect(wb.getSheet("B")!.rowCount).toBe(2);
  });

  it("formats date/datetime columns by pattern into readable strings", async () => {
    const { bytes } = await exportAsStream([
      {
        name: "Dates",
        columns: [
          {
            prop: "d",
            label: "Date",
            format: { type: "date", pattern: "dd/MM/yyyy" },
          },
          { prop: "dt", label: "DateTime", format: { type: "datetime" } },
        ],
        // UTC-constructed so the expected strings hold in every timezone
        // (stream formats UTC components, matching the workbook serial).
        data: [
          {
            d: new Date(Date.UTC(2025, 0, 5, 14, 30)),
            dt: new Date(Date.UTC(2025, 0, 5, 14, 30)),
          },
        ],
      },
    ]);
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("Dates")!;
    // pattern honored; default datetime pattern applied
    expect(ws.cell("A2").value).toBe("05/01/2025");
    expect(ws.cell("B2").value).toBe("2025-01-05 14:30");
  });

  it("writes a spec-correct sharedStrings table (count = total refs, uniqueCount = uniques)", async () => {
    const { bytes } = await exportAsStream([
      {
        name: "SST",
        columns: [
          { prop: "a", label: "A" },
          { prop: "b", label: "B" },
        ],
        data: [
          { a: "dup", b: "x" },
          { a: "dup", b: "x" },
          { a: "uniq", b: "y" },
        ],
      },
    ]);
    // Unzip and inspect the sst directly: per ECMA-376, count is the total
    // number of string-cell references (duplicates included) and uniqueCount
    // the number of distinct strings. 2 header cells + 6 data cells = 8 refs;
    // distinct strings: A, B, dup, uniq, x, y = 6.
    const files = unzipSync(bytes);
    const sst = strFromU8(files["xl/sharedStrings.xml"]);
    expect(sst).toContain('count="8"');
    expect(sst).toContain('uniqueCount="6"');
  });

  it("skips empty cells instead of interning empty strings into sharedStrings", async () => {
    const { bytes } = await exportAsStream([
      {
        name: "S",
        columns: [
          { prop: "a", label: "A" },
          { prop: "b", label: "B" },
        ],
        // `b` missing on every row: previously each missing field interned ""
        // and emitted a shared-string cell reference.
        data: [{ a: "x" }, { a: "y" }],
      },
    ]);
    const files = unzipSync(bytes);
    const sst = strFromU8(files["xl/sharedStrings.xml"]);
    // Only the two headers and the two values are interned; no "" entry.
    expect(sst).toContain('uniqueCount="4"');
    expect(sst).not.toContain('<t xml:space="preserve"></t>');
    // No <c> element is emitted for the empty B-column cells.
    const sheetXml = strFromU8(files["xl/worksheets/sheet1.xml"]);
    expect(sheetXml).not.toContain('r="B2"');
    expect(sheetXml).not.toContain('r="B3"');
    // Round-trip still reads correctly.
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("S")!;
    expect(ws.cell("A2").value).toBe("x");
    expect(ws.rowCount).toBe(3);
  });

  it("clamps the final checkpoint below 1 on whole-thousand row counts", async () => {
    // 整千总数时末个 checkpoint 的值恰为 1，会与 exportExcel 的 terminal 1
    // 重复（onProgress 的 1 必须只出现一次，见 types.ts 契约）——到齐
    // 终点时跳过该 checkpoint，把 1 留给 terminal。
    const progress: number[] = [];
    await exportAsStream(
      [
        {
          name: "S",
          columns: [{ prop: "a", label: "A" }],
          data: Array.from({ length: 2000 }, (_, i) => ({ a: i })),
        },
      ],
      (p) => progress.push(p),
    );
    // 1000 行的 checkpoint 正常上报；2000 行（终点）的不再报。
    expect(progress).toEqual([0.5]);
  });

  it("renders a null data row as an empty row instead of throwing", async () => {
    const { bytes, rowCount } = await exportAsStream([
      {
        name: "S",
        columns: [
          { prop: "a", label: "A" },
          // number/date specs read row[key] directly in displayValue — the
          // null-row guard must cover them too.
          { prop: "b", label: "B", format: { type: "number" } },
          { prop: "c", label: "C", format: { type: "date" } },
        ],
        data: [
          null,
          { a: "x", b: 2, c: new Date(Date.UTC(2025, 0, 5)) },
        ] as unknown as Record<string, unknown>[],
      },
    ]);
    expect(rowCount).toBe(2);
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("S")!;
    expect(ws.rowCount).toBe(3); // header + 2 data rows
    expect(ws.cell("A3").value).toBe("x");
    expect(String(ws.cell("B3").value)).toBe("2");
    expect(ws.cell("C3").value).toBe("2025-01-05");
  });
});
