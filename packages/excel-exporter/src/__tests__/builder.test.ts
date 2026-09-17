import { describe, it, expect } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { WorkbookBuilder } from "../workbook-builder";
import { exportAsStream } from "../streaming-builder";
import { StylePresets } from "../style-presets";
import { readBuffer, makeData } from "./setup";

describe("WorkbookBuilder round-trip", () => {
  it("writes data, headers, styles, freeze, autofilter and merges", async () => {
    const data = makeData(5);
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "Sales",
      freezeRows: 1,
      autoFilter: true,
      merges: [{ row: 0, col: 0, rowspan: 2, colspan: 1 }],
      columns: [
        { prop: "id", label: "ID", width: 10 },
        { prop: "name", label: "Name", width: 18, style: StylePresets.dataRow },
        {
          prop: "amount",
          label: "Amount",
          width: 14,
          style: StylePresets.currency,
        },
        {
          prop: "status",
          label: "Status",
          width: 10,
          format: {
            type: "enum",
            map: { paid: "Paid", pending: "Pending" },
            fallback: "Unknown",
          },
        },
      ],
      data,
    });
    const bytes = await builder.toBuffer();

    // Valid XLSX (ZIP local file header magic)
    expect(bytes.slice(0, 4)).toEqual(new Uint8Array([0x50, 0x4b, 0x03, 0x04]));

    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("Sales")!;
    expect(ws).toBeDefined();

    // Header row present, freeze applied
    expect(ws.cell("A1").value).toBe("ID");
    expect(ws.frozenPane).toEqual({ rows: 1, cols: 0 });

    // Auto-filter covers header..last data row (5 data rows -> A1:D6)
    // autoFilter reads back as AutoFilterData { range: 'A1:D6' }
    expect((ws.autoFilter as { range: string }).range).toBe("A1:D6");

    // Data row 1 values
    expect(String(ws.cell("A2").value)).toBe("0");
    expect(ws.cell("B2").value).toBe("user_0");
    // enum format: paid -> Paid
    expect(ws.cell("D2").value).toBe("Paid");
    expect(ws.cell("D3").value).toBe("Pending");

    // Column style applied: B (dataRow) and C (currency) have non-null styleIndex
    expect(ws.cell("B2").styleIndex).not.toBeNull();
    expect(ws.cell("C2").styleIndex).not.toBeNull();
    // A has no style config -> default (null or 0)
    expect(ws.cell("A2").styleIndex).toBeNull();

    // Style applies to DATA cells only, never the header row (regression guard
    // for the bug where the header cell inherited the column data style).
    expect(ws.cell("B1").styleIndex).toBeNull();
    expect(ws.cell("C1").styleIndex).toBeNull();

    // Merge: A2:A3 (row 0 data-area, rowspan 2 -> rows 2-3 in Excel)
    expect(ws.mergeCells.some((r) => r === "A2:A3")).toBe(true);
  });

  it("still exports columns written with deprecated key/header names", async () => {
    // 2.2.0 兼容承诺：pre-2.2 的 key/header 命名不做迁移也必须继续工作。
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "Legacy",
      columns: [
        { key: "id", header: "ID" },
        { key: "name", header: "Name" },
      ],
      data: [
        { id: 1, name: "Alice" },
        { id: 2, name: "Bob" },
      ],
    });
    const wb = await readBuffer(await builder.toBuffer());
    const ws = wb.getSheet("Legacy")!;
    expect(ws.cell("A1").value).toBe("ID");
    expect(ws.cell("B2").value).toBe("Alice");
    expect(String(ws.cell("A3").value)).toBe("2");
  });

  it("writes a multi-row grouped header with header merges and data offset", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "Grouped",
      freezeRows: 2,
      autoFilter: true,
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
      data: [
        { product: "A", m_qty: 1, m_amt: 2, y_qty: 3, y_amt: 4 },
        { product: "B", m_qty: 5, m_amt: 6, y_qty: 7, y_amt: 8 },
      ],
    });
    const bytes = await builder.toBuffer();
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("Grouped")!;

    // 3 header rows + 2 data rows.
    expect(ws.rowCount).toBe(5);

    // Header values at their merge anchors (top-left of each merged block).
    expect(ws.cell("A1").value).toBe("产品");
    expect(ws.cell("B1").value).toBe("收入情况");
    expect(ws.cell("B2").value).toBe("本月");
    expect(ws.cell("D2").value).toBe("本年累计");
    expect(ws.cell("B3").value).toBe("数量");
    expect(ws.cell("E3").value).toBe("金额");

    // Header merges (leaf spans all header rows, groups span their leaves).
    for (const r of ["A1:A3", "B1:E1", "B2:C2", "D2:E2"]) {
      expect(ws.mergeCells).toContain(r);
    }

    // Data rows start at row 4.
    expect(ws.cell("A4").value).toBe("A");
    expect(ws.cell("E5").value).toBe(8);

    // Auto-filter now anchors on the last header row: A3:E5.
    expect((ws.autoFilter as { range: string }).range).toBe("A3:E5");
  });

  it("applies header styles to merged group/leaf anchors in a grouped header", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "GroupedStyle",
      headerStyle: StylePresets.header,
      columns: [
        {
          prop: "a",
          label: "A",
          headerStyle: StylePresets.danger,
        },
        {
          label: "G",
          headerStyle: StylePresets.currency, // group-level style
          children: [
            { prop: "b", label: "B" },
            { prop: "c", label: "C" },
          ],
        },
      ],
      data: [{ a: 1, b: 2, c: 3 }],
    });
    const bytes = await builder.toBuffer();
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("GroupedStyle")!;

    // Leaf header overrides sheet-level; group header uses its own style; the
    // child leaf B inherits the sheet-level header style.
    expect(ws.cell("A1").styleIndex).not.toBeNull();
    expect(ws.cell("B1").styleIndex).not.toBeNull();
    expect(ws.cell("B2").styleIndex).not.toBeNull();
    // Data cells never take header styles.
    expect(ws.cell("A3").styleIndex).toBeNull();
  });

  it("handles multiple sheets", async () => {
    const builder = await WorkbookBuilder.create();
    builder
      .addSheet({
        name: "S1",
        columns: [{ prop: "a", label: "A" }],
        data: [{ a: 1 }],
      })
      .addSheet({
        name: "S2",
        columns: [{ prop: "b", label: "B" }],
        data: [{ b: 2 }],
      });
    const bytes = await builder.toBuffer();
    const wb = await readBuffer(bytes);
    expect(wb.sheetNames).toEqual(["S1", "S2"]);
    expect(String(wb.getSheet("S1")!.cell("A2").value)).toBe("1");
    expect(String(wb.getSheet("S2")!.cell("A2").value)).toBe("2");
  });

  it("toBlob returns a Blob with xlsx mime type", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "S",
      columns: [{ prop: "x", label: "X" }],
      data: [{ x: "hi" }],
    });
    const blob = await builder.toBlob();
    expect(blob.type).toBe(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    expect(blob.size).toBeGreaterThan(0);
  });

  it("keeps date/number FormatSpec values typed and auto-injects a numFormat", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "Typed",
      columns: [
        {
          prop: "d",
          label: "Date",
          format: { type: "date", pattern: "yyyy/MM/dd" },
        },
        {
          prop: "n",
          label: "Num",
          format: { type: "number", decimals: 2, thousands: true },
        },
        { prop: "plain", label: "Plain" },
      ],
      // 1234.567: full precision must round-trip (was truncated to 1234.57 by
      // toFixed before the fix; decimals=2 now only affects display via numFormat).
      data: [{ d: new Date(2025, 0, 5), n: 1234.567, plain: "x" }],
    });
    const bytes = await builder.toBuffer();
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("Typed")!;

    // Date is an Excel serial (number), not text -- and a style was auto-injected
    // so the cell renders as a date instead of a bare serial number.
    expect(typeof ws.cell("A2").value).toBe("number");
    expect(ws.cell("A2").styleIndex).not.toBeNull();
    // Number is a numeric cell, not text.
    expect(typeof ws.cell("B2").value).toBe("number");
    expect(ws.cell("B2").styleIndex).not.toBeNull();
    // Full precision is preserved in the stored cell (regression guard for the
    // toFixed truncation bug): value is 1234.567, not the display-rounded 1234.57.
    expect(ws.cell("B2").value).toBe(1234.567);
    // Plain column (no format, no style) stays unstyled.
    expect(ws.cell("C2").styleIndex).toBeNull();
  });

  it("normalizes non-primitive values to the same strings as the stream path", async () => {
    // Cross-path contract: a dataset crossing the 50k-row threshold (or
    // degrading to the stream) must keep identical cell content. Before the fix
    // the Workbook path passed raw values to modern-xlsx, which String()ed
    // objects into "[object Object]" and Dates into the localized long form,
    // while the stream path emits JSON / ISO strings via toStr().
    const sheet = {
      name: "Mixed",
      columns: [
        { prop: "obj", label: "Obj" },
        { prop: "d", label: "D" },
        { prop: "big", label: "Big" },
        { prop: "sym", label: "Sym" },
      ],
      data: [
        {
          obj: { a: 1 },
          d: new Date("2025-01-05T00:00:00Z"),
          big: 123n,
          sym: Symbol("s"),
        },
      ],
    };

    const builder = await WorkbookBuilder.create();
    builder.addSheet(sheet);
    const wb = await readBuffer(await builder.toBuffer());
    const ws = wb.getSheet("Mixed")!;
    expect(ws.cell("A2").value).toBe('{"a":1}');
    expect(ws.cell("B2").value).toBe("2025-01-05T00:00:00.000Z");
    expect(ws.cell("C2").value).toBe("123");
    // toStr hardening: JSON.stringify(symbol) is undefined -> String() instead.
    expect(ws.cell("D2").value).toBe("Symbol(s)");

    // The stream path must agree on every one of these cells.
    const { bytes } = await exportAsStream([sheet]);
    const sws = (await readBuffer(bytes)).getSheet("Mixed")!;
    for (const ref of ["A2", "B2", "C2", "D2"]) {
      expect(sws.cell(ref).value).toBe(ws.cell(ref).value);
    }
  });

  it("deduplicates identical styles instead of growing the styles table per cell", async () => {
    // modern-xlsx's StyleBuilder.build appends a fresh font/fill/xf on every
    // call without dedup, so 10 header cells + 10 columns sharing 2 styles
    // previously produced 20 identical records (and 20 distinct styleIndex
    // values). The builder now caches by structural key.
    const builder = await WorkbookBuilder.create();
    const columns = Array.from({ length: 10 }, (_, i) => ({
      prop: `c${i}`,
      label: `C${i}`,
      style: StylePresets.currency,
      headerStyle: StylePresets.header,
    }));
    builder.addSheet({
      name: "Dedup",
      columns,
      data: [Object.fromEntries(columns.map((c) => [c.prop, 1]))],
    });
    const bytes = await builder.toBuffer();

    // Cells sharing one style share one styleIndex.
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("Dedup")!;
    const headerIdx = ws.cell("A1").styleIndex;
    const dataIdx = ws.cell("A2").styleIndex;
    for (let col = 1; col <= 10; col++) {
      const ref = `${String.fromCharCode(64 + col)}1`;
      expect(ws.cell(ref).styleIndex).toBe(headerIdx);
      expect(ws.cell(`${ref.slice(0, -1)}2`).styleIndex).toBe(dataIdx);
    }
    expect(headerIdx).not.toBe(dataIdx);

    // And the serialized styles table stays small: 2 styled xfs (+ modern-xlsx
    // defaults), not one record per styled cell/column.
    const stylesXml = strFromU8(unzipSync(bytes)["xl/styles.xml"]);
    const fontCount = (stylesXml.match(/<font[ >]/g) ?? []).length;
    expect(fontCount).toBeLessThan(6);
  });

  it("renders a null data row as an empty row instead of throwing", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "NullRow",
      columns: [
        { prop: "a", label: "A" },
        { prop: "b", label: "B", format: { type: "number", decimals: 1 } },
      ],
      data: [null, { a: "x", b: 2 }] as unknown as Record<string, unknown>[],
    });
    const bytes = await builder.toBuffer();
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("NullRow")!;
    expect(ws.cell("A3").value).toBe("x");
    expect(String(ws.cell("B3").value)).toBe("2");
  });

  it("accepts a column style whose border object has no sides defined", async () => {
    // buildStyleIndex must not forward an all-undefined border spec to
    // modern-xlsx; the style simply has no border effect.
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "EmptyBorder",
      columns: [{ prop: "a", label: "A", style: { border: {} } }],
      data: [{ a: 1 }],
    });
    const bytes = await builder.toBuffer();
    const wb = await readBuffer(bytes);
    expect(wb.getSheet("EmptyBorder")!.cell("A2").value).toBe(1);
  });
});
