import { describe, it, expect } from "vitest";
import {
  exportEcharts,
  exportTable,
  echartsToSheet,
  tableToSheet,
  WorkbookBuilder,
  StylePresets,
} from "../index";
import { readBuffer } from "./setup";

describe("tableToSheet / exportTable", () => {
  it("rejects circular children with a clear error instead of a stack overflow", () => {
    // 转换阶段的递归 map 先于 flattenColumnTree 的环检测执行：循环引用
    // 若不在 toColumnConfig 拦截，用户拿到的是 RangeError 栈溢出。
    const group: Record<string, unknown> = { title: "G", children: [] };
    (group.children as unknown[]).push(group);
    expect(() =>
      tableToSheet({
        columns: [group as never],
        data: [],
      }),
    ).toThrow(/circular children reference in table columns/);
  });

  it("normalizes Ant Design and Element Plus column naming styles", () => {
    const sheet = tableToSheet({
      sheetName: "Orders",
      columns: [
        { dataIndex: "orderNo", title: "订单号", width: 16 },
        { prop: "customer", label: "客户" },
      ],
      data: [
        { orderNo: "A1", customer: "Alice" },
        { orderNo: "A2", customer: "Bob" },
      ],
    });

    expect(sheet.name).toBe("Orders");
    expect(sheet.columns[0]).toMatchObject({
      prop: "orderNo",
      label: "订单号",
    });
    expect(sheet.columns[1]).toMatchObject({ prop: "customer", label: "客户" });
  });

  it("maps Ant Design grouped columns (children) to a multi-row header", () => {
    const sheet = tableToSheet({
      sheetName: "Grouped",
      columns: [
        { dataIndex: "product", title: "产品" },
        {
          title: "收入情况",
          children: [
            { dataIndex: "m_qty", title: "数量" },
            { dataIndex: "m_amt", title: "金额" },
          ],
        },
      ],
      data: [{ product: "A", m_qty: 1, m_amt: 2 }],
    });

    // Leaf keeps prop/label.
    expect(sheet.columns[0]).toMatchObject({ prop: "product", label: "产品" });
    // Group has label + children, no prop (no data cells).
    expect(sheet.columns[1]).toMatchObject({
      label: "收入情况",
      children: [
        { prop: "m_qty", label: "数量" },
        { prop: "m_amt", label: "金额" },
      ],
    });
    expect(sheet.columns[1].prop).toBeUndefined();
  });

  it("prefers prop/label over legacy key/header; legacy-only columns still work", () => {
    const sheet = tableToSheet({
      columns: [
        { prop: "new", key: "old", label: "New", header: "Old" },
        { key: "legacyOnly", header: "Legacy" },
      ],
      data: [{ new: 1, old: 2, legacyOnly: 3 }],
    });
    expect(sheet.columns[0]).toMatchObject({ prop: "new", label: "New" });
    // 旧名（pre-2.2 命名）单独使用是受支持的兼容路径，输出统一为新名。
    expect(sheet.columns[1]).toMatchObject({
      prop: "legacyOnly",
      label: "Legacy",
    });
  });

  it("exports a real xlsx from a common table data shape", async () => {
    const result = await exportTable({
      filename: "table-adapter",
      download: false,
      mode: "main",
      sheetName: "表格",
      columns: [
        { dataIndex: "orderNo", title: "订单号" },
        { dataIndex: "customer", title: "客户" },
        {
          dataIndex: "amount",
          title: "金额",
          format: { type: "number", decimals: 2 },
        },
      ],
      data: [
        { orderNo: "A1", customer: "Alice", amount: 1234.5 },
        { orderNo: "A2", customer: "Bob", amount: 88 },
      ],
    });

    expect(result.success).toBe(true);
    const wb = await readBuffer(
      new Uint8Array(await result.blob!.arrayBuffer()),
    );
    const ws = wb.getSheet("表格")!;
    expect(ws.cell("A1").value).toBe("订单号");
    expect(ws.cell("C3").value).toBe(88);
  });
});

describe("echartsToSheet / exportEcharts", () => {
  it("exports category + multiple series as a wide table", () => {
    const sheet = echartsToSheet({
      option: {
        xAxis: { type: "category", data: ["一月", "二月", "三月"] },
        series: [
          { name: "销售额", type: "line", data: [120, 200, 150] },
          { name: "利润", type: "line", data: [30, 60, 45] },
        ],
      },
    });

    expect(sheet.columns.map((c) => c.label)).toEqual([
      "类目",
      "销售额",
      "利润",
    ]);
    expect(sheet.data[1]).toEqual({
      类目: "二月",
      __series_0: 200,
      __series_1: 60,
    });
  });

  it("exports category + multiple series as a long table", () => {
    const sheet = echartsToSheet({
      layout: "long",
      option: {
        xAxis: { data: ["A", "B"] },
        series: [{ name: "S1", data: [1, 2] }],
      },
    });

    expect(sheet.columns.map((c) => c.label)).toEqual(["系列", "类目", "数值"]);
    expect(sheet.data).toHaveLength(2);
    expect(sheet.data[1]).toEqual({ 系列: "S1", 类目: "B", 数值: 2 });
  });

  it("exports pie-like name/value data", () => {
    const sheet = echartsToSheet({
      option: {
        series: [
          {
            name: "占比",
            type: "pie",
            data: [
              { name: "A", value: 10 },
              { name: "B", value: 20 },
            ],
          },
        ],
      },
    });

    expect(sheet.columns.map((c) => c.label)).toEqual(["系列", "名称", "数值"]);
    expect(sheet.data).toEqual([
      { 系列: "占比", 名称: "A", 数值: 10 },
      { 系列: "占比", 名称: "B", 数值: 20 },
    ]);
  });

  it("exports scatter-like coordinate pairs", () => {
    const sheet = echartsToSheet({
      option: {
        series: [
          {
            name: "点",
            type: "scatter",
            data: [
              [1, 2],
              [3, 4],
            ],
          },
        ],
      },
    });

    expect(sheet.columns.map((c) => c.label)).toEqual(["系列", "X", "Y"]);
    expect(sheet.data).toEqual([
      { 系列: "点", X: 1, Y: 2 },
      { 系列: "点", X: 3, Y: 4 },
    ]);
  });

  it("exports object-form scatter datums ({ value: [x, y] }) like bare pairs", () => {
    // Both spellings are legal ECharts scatter data; the object form used to
    // fall through to the name/value branch and stringify the pair as "[1,2]".
    const sheet = echartsToSheet({
      option: {
        series: [
          {
            name: "点",
            type: "scatter",
            data: [{ value: [1, 2], name: "ignored" }, { value: [3, 4] }],
          },
        ],
      },
    });

    expect(sheet.columns.map((c) => c.label)).toEqual(["系列", "X", "Y"]);
    expect(sheet.data).toEqual([
      { 系列: "点", X: 1, Y: 2 },
      { 系列: "点", X: 3, Y: 4 },
    ]);
  });

  it("mixes the two scatter spellings in one series, and with other series", () => {
    const sheet = echartsToSheet({
      option: {
        series: [
          { name: "a", data: [[1, 2], { value: [3, 4] }] },
          { name: "b", data: [{ value: [5, 6] }] },
        ],
      },
    });

    expect(sheet.columns.map((c) => c.label)).toEqual(["系列", "X", "Y"]);
    expect(sheet.data).toEqual([
      { 系列: "a", X: 1, Y: 2 },
      { 系列: "a", X: 3, Y: 4 },
      { 系列: "b", X: 5, Y: 6 },
    ]);
  });

  it("still rejects scatter data (either spelling) mixed with name/value data", () => {
    expect(() =>
      echartsToSheet({
        option: {
          series: [{ name: "s", data: [[1, 2], { name: "n", value: 3 }] }],
        },
      }),
    ).toThrow(/mixing scatter coordinate data/);

    expect(() =>
      echartsToSheet({
        option: {
          series: [{ name: "s", data: [{ value: [1, 2] }, "plain"] }],
        },
      }),
    ).toThrow(/mixing scatter coordinate data/);
  });

  it("rejects dataset mode instead of guessing", () => {
    expect(() =>
      echartsToSheet({
        option: {
          dataset: { source: [] },
          series: [],
        },
      }),
    ).toThrow(/dataset mode is not supported/);
  });

  it("exports a real xlsx from a simple ECharts option", async () => {
    const result = await exportEcharts({
      filename: "echarts-adapter",
      download: false,
      mode: "main",
      option: {
        xAxis: { data: ["一月", "二月"] },
        series: [{ name: "销售额", data: [120, 200] }],
      },
    });

    expect(result.success).toBe(true);
    const wb = await readBuffer(
      new Uint8Array(await result.blob!.arrayBuffer()),
    );
    const ws = wb.getSheet("图表数据")!;
    expect(ws.cell("A1").value).toBe("类目");
    expect(ws.cell("B3").value).toBe(200);
  });
});

describe("headerStyle", () => {
  it("applies sheet-level and column-level header styles without styling data cells", async () => {
    const builder = await WorkbookBuilder.create();
    builder.addSheet({
      name: "StyledHeader",
      headerStyle: StylePresets.header,
      columns: [
        { prop: "a", label: "A", style: StylePresets.dataRow },
        {
          prop: "b",
          label: "B",
          headerStyle: StylePresets.danger,
          style: StylePresets.dataRow,
        },
      ],
      data: [
        { a: 1, b: 2 },
        { a: 3, b: 4 },
      ],
    });

    const bytes = await builder.toBuffer();
    const wb = await readBuffer(bytes);
    const ws = wb.getSheet("StyledHeader")!;

    expect(ws.cell("A1").styleIndex).not.toBeNull();
    expect(ws.cell("B1").styleIndex).not.toBeNull();
    expect(ws.cell("A2").styleIndex).not.toBeNull();
    expect(ws.cell("B2").styleIndex).not.toBeNull();
  });
});

describe("echartsToSheet header collisions", () => {
  // Long/item layouts key data rows by the header text itself, so duplicated
  // headers would silently overwrite each other's column.
  it("rejects duplicate headers in category long layout", () => {
    expect(() =>
      echartsToSheet({
        layout: "long",
        categoryHeader: "数值",
        option: {
          xAxis: { data: ["A"] },
          series: [{ name: "S1", data: [1] }],
        },
      }),
    ).toThrow(/duplicate header "数值"/);
  });

  it("rejects a seriesHeader colliding with X/Y in scatter layout", () => {
    expect(() =>
      echartsToSheet({
        seriesHeader: "X",
        option: { series: [{ data: [[1, 2]] }] },
      }),
    ).toThrow(/duplicate header "X"/);
  });

  it("rejects duplicate headers in name/value layout", () => {
    expect(() =>
      echartsToSheet({
        nameHeader: "数值",
        option: { series: [{ data: [{ name: "A", value: 1 }] }] },
      }),
    ).toThrow(/duplicate header "数值"/);
  });

  it("rejects a categoryHeader colliding with the internal __series_N keys (wide layout)", () => {
    // Wide layout keys rows by categoryHeader plus internal __series_N keys;
    // a collision would silently overwrite the category column with series data.
    expect(() =>
      echartsToSheet({
        categoryHeader: "__series_0",
        option: {
          xAxis: { data: ["A"] },
          series: [{ name: "S1", data: [1] }],
        },
      }),
    ).toThrow(/collides with the internal series keys/);
  });
});

describe("dataStyle / indexColumn passthrough (exportTable)", () => {
  it("passes sheet-level dataStyle and indexColumn into SheetConfig", () => {
    const sheet = tableToSheet({
      sheetName: "T",
      dataStyle: StylePresets.bordered,
      indexColumn: { label: "序号", width: 8 },
      columns: [{ dataIndex: "name", title: "名称" }],
      data: [{ name: "a" }],
    });
    expect(sheet.dataStyle).toBe(StylePresets.bordered);
    expect(sheet.indexColumn).toEqual({ label: "序号", width: 8 });
  });

  it("keeps both fields absent when the caller does not set them", () => {
    const sheet = tableToSheet({
      columns: [{ prop: "name", label: "名称" }],
      data: [{ name: "a" }],
    });
    expect(sheet.dataStyle).toBeUndefined();
    expect(sheet.indexColumn).toBeUndefined();
  });
});
