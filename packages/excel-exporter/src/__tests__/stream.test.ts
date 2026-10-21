import { describe, it, expect } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { dateToSerial } from "@marcusok/xlsx-core";
import { exportAsStream } from "../streaming-builder";
import { exportExcel } from "../index";
import type { SheetConfig } from "../types";
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

  it("stores date/datetime as serials with a numFmt style", async () => {
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
        // UTC-constructed: dateToSerial uses UTC components, same as the
        // Workbook path.
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
    // 存储层是真日期序列号，不再是格式化文本；显示格式由 numFmt 承担。
    expect(Number(ws.cell("A2").value)).toBeCloseTo(
      dateToSerial(new Date(Date.UTC(2025, 0, 5, 14, 30))),
      10,
    );
    expect(Number(ws.cell("B2").value)).toBeCloseTo(
      dateToSerial(new Date(Date.UTC(2025, 0, 5, 14, 30))),
      10,
    );
    const files = unzipSync(bytes);
    const styles = strFromU8(files["xl/styles.xml"]);
    expect(styles).toContain('formatCode="dd/MM/yyyy"');
    expect(styles).toContain('formatCode="yyyy-MM-dd HH:mm"');
    const sheetXml = strFromU8(files["xl/worksheets/sheet1.xml"]);
    expect(sheetXml).toMatch(/<c r="A2" s="\d+"/);
    expect(sheetXml).toMatch(/<c r="B2" s="\d+"/);
  });

  it("keeps date/number cell values identical across Workbook and stream paths", async () => {
    const sheet: SheetConfig = {
      name: "Cross",
      columns: [
        {
          prop: "d",
          label: "D",
          format: { type: "date", pattern: "yyyy-MM-dd" },
        },
        {
          prop: "n",
          label: "N",
          format: { type: "number", decimals: 2, thousands: true },
        },
      ],
      data: [
        {
          d: new Date(Date.UTC(2025, 0, 5)),
          n: 1234.567,
        },
      ],
    };
    const main = await exportExcel({
      filename: "cross-main",
      download: false,
      mode: "main",
      sheets: [sheet],
    });
    const stream = await exportExcel({
      filename: "cross-stream",
      download: false,
      mode: "stream",
      sheets: [sheet],
    });
    expect(main.success).toBe(true);
    expect(stream.success).toBe(true);
    const mainWs = (
      await readBuffer(new Uint8Array(await main.blob!.arrayBuffer()))
    ).getSheet("Cross")!;
    const streamWs = (
      await readBuffer(new Uint8Array(await stream.blob!.arrayBuffer()))
    ).getSheet("Cross")!;
    // 存储值一致：日期都是序列号、金额都保留完整精度（显示由 numFmt 负责）。
    expect(Number(streamWs.cell("A2").value)).toBeCloseTo(
      Number(mainWs.cell("A2").value),
      10,
    );
    expect(streamWs.cell("B2").value).toBe(mainWs.cell("B2").value);
    expect(streamWs.cell("B2").value).toBe(1234.567);
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
    expect(Number(ws.cell("C3").value)).toBeCloseTo(
      dateToSerial(new Date(Date.UTC(2025, 0, 5))),
      10,
    );
  });
});
