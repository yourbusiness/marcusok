import { describe, it, expect } from "vitest";
import { columnProp, flattenColumnTree, a1Range } from "../column-tree";
import type { ColumnConfig } from "../types";

/** 3-level grouped label: 产品(leaf) | 收入情况(本月/本年累计). */
const groupedColumns: ColumnConfig[] = [
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
];

describe("flattenColumnTree", () => {
  it("keeps flat columns byte-compatible: H=1, no header merges, same leaf order", () => {
    const flat = flattenColumnTree([
      { prop: "a", label: "A" },
      { prop: "b", label: "B" },
    ]);
    expect(flat.headerRowCount).toBe(1);
    expect(flat.leaves.map((l) => l.prop)).toEqual(["a", "b"]);
    expect(flat.headerGrid).toEqual([["A", "B"]]);
    // Single-cell spans must not produce merges, so flat output is unchanged.
    expect(flat.headerMerges).toHaveLength(0);
    expect(flat.headerCells).toHaveLength(2);
  });

  it("flattens a 3-level grouped header into leaves + merges", () => {
    const t = flattenColumnTree(groupedColumns);
    expect(t.headerRowCount).toBe(3);
    expect(t.leaves.map((l) => l.prop)).toEqual([
      "product",
      "m_qty",
      "m_amt",
      "y_qty",
      "y_amt",
    ]);

    // 产品 spans the 3 header rows in column 0.
    expect(t.headerGrid[0][0]).toBe("产品");
    expect(t.headerGrid[1][0]).toBeNull();
    expect(t.headerGrid[2][0]).toBeNull();
    // Group headers at their own depth.
    expect(t.headerGrid[0][1]).toBe("收入情况");
    expect(t.headerGrid[1][1]).toBe("本月");
    expect(t.headerGrid[1][3]).toBe("本年累计");
    // Leaf headers on the bottom row.
    expect(t.headerGrid[2][1]).toBe("数量");
    expect(t.headerGrid[2][2]).toBe("金额");
    expect(t.headerGrid[2][3]).toBe("数量");
    expect(t.headerGrid[2][4]).toBe("金额");

    const merges = t.headerMerges.map((m) =>
      a1Range(m.row, m.col, m.rowSpan, m.colSpan),
    );
    expect(merges).toEqual(["A1:A3", "B1:E1", "B2:C2", "D2:E2"]);
  });

  it("handles mixed depth: a depth-0 leaf next to a 1-level group", () => {
    const t = flattenColumnTree([
      { prop: "a", label: "A" },
      {
        label: "G",
        children: [
          { prop: "b", label: "B" },
          { prop: "c", label: "C" },
        ],
      },
    ]);
    expect(t.headerRowCount).toBe(2);
    // Leaf A spans both header rows.
    expect(
      t.headerMerges.map((m) => a1Range(m.row, m.col, m.rowSpan, m.colSpan)),
    ).toEqual(["A1:A2", "B1:C1"]);
    expect(t.headerGrid).toEqual([
      ["A", "G", null],
      [null, "B", "C"],
    ]);
    expect(t.leaves.map((l) => l.prop)).toEqual(["a", "b", "c"]);
  });

  it("treats children: [] as a leaf", () => {
    const t = flattenColumnTree([{ prop: "a", label: "A", children: [] }]);
    expect(t.headerRowCount).toBe(1);
    expect(t.leaves.map((l) => l.prop)).toEqual(["a"]);
  });

  it("throws when a leaf column has no usable prop", () => {
    expect(() => flattenColumnTree([{ label: "no prop" }])).toThrow(
      /must have a non-empty string prop/,
    );
  });

  it("reads deprecated key/header aliases; prop/label win when both exist", () => {
    // 2.2.0 前的旧字段名仍可用：仅提供旧名时导出行为不变。
    const legacy = flattenColumnTree([
      { key: "a", header: "A" },
      { header: "G", children: [{ key: "b", header: "B" }] },
    ]);
    expect(legacy.leaves.map(columnProp)).toEqual(["a", "b"]);
    expect(legacy.headerGrid).toEqual([
      ["A", "G"],
      [null, "B"],
    ]);
    // 新旧同给时新名优先，旧名被忽略。
    const both = flattenColumnTree([
      { key: "old", header: "Old", prop: "new", label: "New" },
    ]);
    expect(both.leaves.map(columnProp)).toEqual(["new"]);
    expect(both.headerGrid).toEqual([["New"]]);
  });

  it("throws when a column has neither label nor legacy header", () => {
    expect(() => flattenColumnTree([{ prop: "a" }])).toThrow(
      /must have a non-empty label/,
    );
  });

  it("throws on circular children references instead of overflowing", () => {
    const a: ColumnConfig = { label: "A", children: [] };
    const b: ColumnConfig = { label: "B", children: [a] };
    a.children!.push(b);
    expect(() => flattenColumnTree([a])).toThrow(/circular children/);
  });

  it("throws when the same column object is reused (diamond, not a cycle)", () => {
    const leaf: ColumnConfig = { prop: "a", label: "A" };
    // Hung under two parents: previously passed the path-based cycle check
    // and was walked twice, silently emitting duplicate data columns.
    expect(() =>
      flattenColumnTree([
        { label: "G1", children: [leaf] },
        { label: "G2", children: [leaf] },
      ]),
    ).toThrow(/reused/);
  });

  it("a1Range renders 0-based ranges as A1 refs", () => {
    expect(a1Range(0, 0, 1, 1)).toBe("A1:A1");
    expect(a1Range(0, 0, 3, 1)).toBe("A1:A3");
    expect(a1Range(2, 1, 1, 4)).toBe("B3:E3");
    expect(a1Range(1, 26, 1, 2)).toBe("AA2:AB2");
  });
});
