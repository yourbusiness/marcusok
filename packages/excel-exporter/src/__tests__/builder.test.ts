import { describe, it, expect } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { WorkbookBuilder } from "../workbook-builder";
import { exportAsStream } from "../streaming-builder";
import { StylePresets } from "../style-presets";
import { applyIndexColumn } from "../sheet-normalize";
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
    // A has no column style -> it still carries the library base style (centered),
    // which is structurally distinct from B/C's merged styles.
    const aIdx = ws.cell("A2").styleIndex;
    expect(aIdx).not.toBeNull();
    expect(aIdx).not.toBe(ws.cell("B2").styleIndex);
    expect(aIdx).not.toBe(ws.cell("C2").styleIndex);

    // Style applies to DATA cells only, never the header row (regression guard
    // for the bug where the header cell inherited the column data style). Header
    // cells now carry the base style as well, so compare indices instead of
    // asserting null.
    expect(ws.cell("B1").styleIndex).not.toBe(ws.cell("B2").styleIndex);
    expect(ws.cell("C1").styleIndex).not.toBe(ws.cell("C2").styleIndex);

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
    // Data cells never take header styles: the data cell carries only the base
    // style, so its index differs from the leaf header's own style.
    expect(ws.cell("A3").styleIndex).not.toBe(ws.cell("A1").styleIndex);
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
    // Plain column (no format, no style) gets no auto-injected numFormat -- only
    // the base style, so its index differs from the two typed columns.
    expect(ws.cell("C2").styleIndex).not.toBeNull();
    expect(ws.cell("C2").styleIndex).not.toBe(ws.cell("A2").styleIndex);
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

describe("BaseCellStyle (library default centering)", () => {
  it("centers cells that declare no alignment, and yields to explicit values", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "Base",
      columns: [
        { prop: "plain", label: "Plain" }, // 无任何样式 -> 纯基底
        // 显式左对齐：水平覆盖基底，垂直仍由基底兜底
        {
          prop: "left",
          label: "Left",
          style: { alignment: { horizontal: "left" } },
        },
        // 预设右对齐（currency）：同上，且带 numFormat
        { prop: "cur", label: "Cur", style: StylePresets.currency },
      ],
      data: [{ plain: "x", left: "y", cur: 1.5 }],
    });
    const bytes = await builder.toBuffer();
    const stylesXml = strFromU8(unzipSync(bytes)["xl/styles.xml"]);

    // 基底：未声明对齐的单元格落到水平 + 垂直居中
    expect(stylesXml).toContain(
      '<alignment horizontal="center" vertical="center"/>',
    );
    // 显式声明优先：水平被覆盖，垂直仍由基底补上
    expect(stylesXml).toContain(
      '<alignment horizontal="left" vertical="center"/>',
    );
    expect(stylesXml).toContain(
      '<alignment horizontal="right" vertical="center"/>',
    );

    // 表头同样吃基底（本表未配 headerStyle）：各列表头共享同一基底样式索引
    const ws = (await readBuffer(bytes)).getSheet("Base")!;
    expect(ws.cell("A1").styleIndex).not.toBeNull();
    expect(ws.cell("B1").styleIndex).toBe(ws.cell("A1").styleIndex);
  });

  it("dataStyle overrides the base alignment for every data cell", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "Override",
      // 调用方要 Excel 原生左对齐时的正路：用 dataStyle 覆盖基底
      dataStyle: { alignment: { horizontal: "left" } },
      columns: [{ prop: "a", label: "A" }],
      data: [{ a: 1 }],
    });
    const stylesXml = strFromU8(
      unzipSync(await builder.toBuffer())["xl/styles.xml"],
    );
    expect(stylesXml).toContain(
      '<alignment horizontal="left" vertical="center"/>',
    );
  });
});

describe("dataStyle (sheet-level base style)", () => {
  it("applies to every data cell and never to header cells", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "DataStyle",
      dataStyle: StylePresets.bordered,
      columns: [
        { prop: "a", label: "A" },
        { prop: "b", label: "B" },
      ],
      data: [
        { a: 1, b: 2 },
        { a: 3, b: 4 },
      ],
    });
    const wb = await readBuffer(await builder.toBuffer());
    const ws = wb.getSheet("DataStyle")!;
    for (const ref of ["A2", "B2", "A3", "B3"]) {
      expect(ws.cell(ref).styleIndex).not.toBeNull();
    }
    // dataStyle 只作用于数据区：表头不被波及（与 headerStyle 的分工一致）。
    // 表头现在也带库级基底，故与数据格的合并结果比较，而非断言为 null。
    expect(ws.cell("A1").styleIndex).not.toBe(ws.cell("A2").styleIndex);
    expect(ws.cell("B1").styleIndex).not.toBe(ws.cell("B2").styleIndex);
  });

  it("deep-merges: a column style keeps the base border while overriding alignment", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "Merged",
      dataStyle: StylePresets.bordered,
      columns: [
        // 列级只改对齐与数字格式：边框应继承表级基底（字段级合并，而非整体替换）
        { prop: "a", label: "A", style: StylePresets.currency },
        { prop: "b", label: "B" },
      ],
      data: [
        { a: 1.5, b: 2 },
        { a: 3.5, b: 4 },
      ],
    });
    const wb = await readBuffer(await builder.toBuffer());
    const ws = wb.getSheet("Merged")!;
    const aIdx = ws.cell("A2").styleIndex;
    const bIdx = ws.cell("B2").styleIndex;
    expect(aIdx).not.toBeNull();
    expect(bIdx).not.toBeNull();
    // A 列合并了两种来源，B 列只有基底——结构不同必然是不同的样式索引
    expect(aIdx).not.toBe(bIdx);
    // 同列内所有数据单元格共享一个合并结果（styleIndexCache 以最终合并为 key）
    expect(ws.cell("A3").styleIndex).toBe(aIdx);
    expect(ws.cell("B3").styleIndex).toBe(bIdx);
  });

  it("auto numFormat injection still wins over a dataStyle numFormat", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "NumFmt",
      dataStyle: { numFormat: "#,##0" },
      columns: [
        // FormatSpec(number, decimals 2) 自动注入的 numFormat 应优先于表级装饰性格式
        { prop: "a", label: "A", format: { type: "number", decimals: 2 } },
      ],
      data: [{ a: 1 }],
    });
    const wb = await readBuffer(await builder.toBuffer());
    const ws = wb.getSheet("NumFmt")!;
    expect(ws.cell("A2").styleIndex).not.toBeNull();
  });
});

describe("indexColumn (workbook path)", () => {
  it("generates row numbers from the row index, never reading data", async () => {
    // WorkbookBuilder 是低级直连路径，不经过 exportExcel 的入口归一化；
    // 用 applyIndexColumn 显式展开（等价于入口对 indexColumn 的处理）。
    const builder = await WorkbookBuilder.create();
    builder.addSheet(
      applyIndexColumn({
        name: "Idx",
        indexColumn: { label: "序号", start: 1 },
        columns: [
          { prop: "name", label: "名称" },
          { prop: "amount", label: "金额" },
        ],
        // 数据行不带任何序号字段：值必须来自行号而非 data
        data: [
          { name: "a", amount: 1 },
          { name: "b", amount: 2 },
          { name: "c", amount: 3 },
        ],
      }),
    );
    const wb = await readBuffer(await builder.toBuffer());
    const ws = wb.getSheet("Idx")!;
    expect(ws.cell("A1").value).toBe("序号");
    expect(ws.cell("B1").value).toBe("名称");
    expect(String(ws.cell("A2").value)).toBe("1");
    expect(String(ws.cell("A3").value)).toBe("2");
    expect(String(ws.cell("A4").value)).toBe("3");
    expect(ws.cell("B2").value).toBe("a");
  });

  it("honors a custom start and shifts user merges right by one column", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet(
      applyIndexColumn({
        name: "IdxStart",
        indexColumn: { start: 10 },
        merges: [{ row: 0, col: 0, rowspan: 2, colspan: 1 }],
        columns: [
          { prop: "name", label: "名称" },
          { prop: "amount", label: "金额" },
        ],
        data: [
          { name: "a", amount: 1 },
          { name: "b", amount: 2 },
        ],
      }),
    );
    const wb = await readBuffer(await builder.toBuffer());
    const ws = wb.getSheet("IdxStart")!;
    expect(String(ws.cell("A2").value)).toBe("10");
    expect(String(ws.cell("A3").value)).toBe("11");
    // 用户 merge 原指向第 0 数据列（name），偏移后应落在 name 列 B2:B3
    expect(ws.mergeCells.some((r) => r === "B2:B3")).toBe(true);
  });

  it("stream path writes the same index numbers (cross-path consistency)", async () => {
    const { bytes } = await exportAsStream([
      applyIndexColumn({
        name: "IdxStream",
        indexColumn: true,
        columns: [
          { prop: "name", label: "名称" },
          { prop: "amount", label: "金额" },
        ],
        data: [
          { name: "a", amount: 1 },
          { name: "b", amount: 2 },
        ],
      }),
    ]);
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("IdxStream")!;
    expect(ws.cell("A1").value).toBe("序号");
    expect(String(ws.cell("A2").value)).toBe("1");
    expect(String(ws.cell("A3").value)).toBe("2");
    expect(ws.cell("B2").value).toBe("a");
  });
});

describe("indexColumn with a multi-row grouped header", () => {
  it("spans the index header vertically across all header rows", async () => {
    // 序号列是顶层叶子列：列树最大深度为 2 时，它的表头须纵向跨满 2 行
    // （与其他叶子列的既有行为一致，由 flattenColumnTree 自动生成）。
    const builder = await WorkbookBuilder.create();
    builder.addSheet(
      applyIndexColumn({
        name: "IdxGrouped",
        indexColumn: { label: "序号" },
        columns: [
          {
            label: "客户信息",
            children: [
              { prop: "name", label: "名称" },
              { prop: "city", label: "城市" },
            ],
          },
        ],
        data: [
          { name: "a", city: "x" },
          { name: "b", city: "y" },
        ],
      }),
    );
    const wb = await readBuffer(await builder.toBuffer());
    const ws = wb.getSheet("IdxGrouped")!;
    expect(ws.cell("A1").value).toBe("序号");
    expect(ws.cell("B1").value).toBe("客户信息");
    expect(ws.cell("B2").value).toBe("名称");
    expect(ws.cell("C2").value).toBe("城市");
    expect(String(ws.cell("A3").value)).toBe("1");
    expect(ws.cell("B3").value).toBe("a");
    // 序号表头纵向合并 A1:A2；分组表头横向合并 B1:C1
    expect(ws.mergeCells.some((r) => r === "A1:A2")).toBe(true);
    expect(ws.mergeCells.some((r) => r === "B1:C1")).toBe(true);
  });

  it("indexColumnStart tolerates a stray null without throwing mid-build", async () => {
    // JS 调用方可能传 indexColumn: null（构建器直连路径不过 validateInput）：
    // 归一化跳过它、不注入虚拟列，indexColumnStart 也不能因此抛错。
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "NullIdx",
      indexColumn: null as unknown as boolean,
      columns: [{ prop: "name", label: "名称" }],
      data: [{ name: "a" }],
    });
    const wb = await readBuffer(await builder.toBuffer());
    const ws = wb.getSheet("NullIdx")!;
    expect(ws.cell("A1").value).toBe("名称");
    expect(ws.cell("A2").value).toBe("a");
  });
});
