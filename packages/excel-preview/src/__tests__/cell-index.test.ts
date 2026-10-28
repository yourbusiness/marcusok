/**
 * cell-index 单元测试：构建语义（0-based 行号、乱序排序、同列去重保末位）
 * 与二分查询边界（点查命中/越界、区间开闭、空行/空区间）。
 *
 * 渲染层的真实窗口循环由 render.test.ts 覆盖（经 SheetRenderer 全链路），
 * 这里只钉索引模块自身的契约——render 测试的模型 cells 恒有序无重复，
 * needSort/needDedup 两条防御路径只有本文件直接触发。
 */
import { describe, it, expect } from "vitest";
import {
  buildRowIndex,
  cellAt,
  forEachCellInRange,
} from "../render/cell-index";
import type { PreviewCell, PreviewRow } from "../types";

function cell(col: number, value = String(col)): PreviewCell {
  return { col, type: "number", value, styleIndex: null };
}

function row(index: number, cells: PreviewCell[]): PreviewRow {
  return { index, height: null, hidden: false, cells };
}

describe("cell-index", () => {
  it("indexes rows by 0-based row number (model is 1-based)", () => {
    const index = buildRowIndex([row(1, [cell(0)]), row(3, [cell(5)])]);
    expect(index.size).toBe(2);
    expect(index.has(0)).toBe(true);
    expect(index.has(2)).toBe(true);
    expect(index.has(1)).toBe(false); // 无内容的第 2 行不在模型里
  });

  it("sorts unsorted cells and dedupes duplicate cols keeping the last", () => {
    // 同列重复 + 乱序同时出现：排序后去重保末位（与旧 Map.set 的覆盖
    // 语义一致——后写胜出）
    const index = buildRowIndex([
      row(1, [cell(5, "a"), cell(2, "b"), cell(5, "c"), cell(0, "d")]),
    ]);
    const cells = index.get(0)!;
    expect(cells.map((c) => c.col)).toEqual([0, 2, 5]);
    expect(cells[2].value).toBe("c");
  });

  it("does not mutate the model's cells array", () => {
    // 模型不可变是渲染层的前提（同一份模型可被多实例共享）：
    // 排序/去重必须发生在副本上
    const cells = [cell(3), cell(1)];
    buildRowIndex([row(1, cells)]);
    expect(cells.map((c) => c.col)).toEqual([3, 1]);
  });

  it("cellAt: hit, gap miss, past-end miss, and missing row", () => {
    const index = buildRowIndex([row(1, [cell(3), cell(7)])]);
    expect(cellAt(index.get(0), 3)?.value).toBe("3");
    expect(cellAt(index.get(0), 7)?.value).toBe("7");
    expect(cellAt(index.get(0), 4)).toBeUndefined(); // 列缝隙
    expect(cellAt(index.get(0), 8)).toBeUndefined(); // 越过末格
    expect(cellAt(index.get(0), -1)).toBeUndefined();
    expect(cellAt(undefined, 1)).toBeUndefined();
  });

  it("forEachCellInRange iterates [colStart, colEnd) in column order", () => {
    const index = buildRowIndex([row(1, [cell(1), cell(2), cell(3), cell(4)])]);
    const seen: number[] = [];
    forEachCellInRange(index.get(0), 2, 4, (c) => seen.push(c.col));
    expect(seen).toEqual([2, 3]); // 右界开区间
    // 空/反向区间与缺失行都不回调
    forEachCellInRange(index.get(0), 3, 3, () => seen.push(-1));
    forEachCellInRange(index.get(0), 4, 2, () => seen.push(-2));
    forEachCellInRange(undefined, 0, 9, () => seen.push(-3));
    expect(seen).toEqual([2, 3]);
  });

  it("range starting before the first cell clamps to the first hit", () => {
    const index = buildRowIndex([row(1, [cell(100), cell(200)])]);
    const seen: number[] = [];
    forEachCellInRange(index.get(0), 0, 300, (c) => seen.push(c.col));
    expect(seen).toEqual([100, 200]);
  });

  it("empty rows produce empty arrays (kept in the index, like the old Map)", () => {
    // 空行（自定义行高的无格行）保留占位：与旧 Map.set(r, new Map()) 一致
    const index = buildRowIndex([row(1, [])]);
    expect(index.get(0)).toEqual([]);
  });
});
