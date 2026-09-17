import { describe, it, expect } from "vitest";
import {
  applyIndexColumn,
  indexColumnStart,
  INDEX_PROP,
} from "../sheet-normalize";
import type { SheetConfig } from "../types";

const baseSheet = (over: Partial<SheetConfig> = {}): SheetConfig => ({
  name: "S",
  columns: [
    { prop: "name", label: "名称" },
    { prop: "amount", label: "金额" },
  ],
  data: [
    { name: "a", amount: 1 },
    { name: "b", amount: 2 },
  ],
  ...over,
});

describe("applyIndexColumn", () => {
  it("returns the sheet untouched when indexColumn is absent or false", () => {
    const sheet = baseSheet();
    expect(applyIndexColumn(sheet)).toBe(sheet);
    const falsy = baseSheet({ indexColumn: false });
    expect(applyIndexColumn(falsy)).toBe(falsy);
  });

  it("expands the shorthand true into defaults: 序号 header, width 6, first leaf", () => {
    const sheet = baseSheet({ indexColumn: true });
    const out = applyIndexColumn(sheet);
    expect(out.columns).toHaveLength(3);
    expect(out.columns[0]).toEqual({
      prop: INDEX_PROP,
      label: "序号",
      width: 6,
    });
    expect(out.columns[1]).toEqual({ prop: "name", label: "名称" });
    // 原 sheet 不被就地修改（入口归一化后用户对象保持原样）
    expect(sheet.columns).toHaveLength(2);
  });

  it("applies custom label / width / start / styles; width 0 survives", () => {
    const out = applyIndexColumn(
      baseSheet({
        indexColumn: {
          label: "序号",
          width: 0,
          style: { alignment: { horizontal: "center" } },
          headerStyle: { font: { bold: true } },
        },
      }),
    );
    // width 0 是合法的"隐藏列"，不能被默认值 6 吞掉（?? 而非 ||）
    expect(out.columns[0]).toEqual({
      prop: INDEX_PROP,
      label: "序号",
      width: 0,
      style: { alignment: { horizontal: "center" } },
      headerStyle: { font: { bold: true } },
    });
    // start 是构建期参数，不落在列配置上
    expect("start" in out.columns[0]).toBe(false);
  });

  it("shifts data-area merges one column right so they keep their targets", () => {
    const out = applyIndexColumn(
      baseSheet({
        indexColumn: true,
        merges: [
          { row: 0, col: 0, rowspan: 2, colspan: 1 },
          { row: 1, col: 1, rowspan: 1, colspan: 1 },
        ],
      }),
    );
    expect(out.merges).toEqual([
      { row: 0, col: 1, rowspan: 2, colspan: 1 },
      { row: 1, col: 2, rowspan: 1, colspan: 1 },
    ]);
  });

  it("drops merges handling when the sheet has none (undefined stays undefined)", () => {
    const out = applyIndexColumn(baseSheet({ indexColumn: true }));
    expect(out.merges).toBeUndefined();
  });

  it("rejects a user column already claiming the reserved prop", () => {
    expect(() =>
      applyIndexColumn(
        baseSheet({
          indexColumn: true,
          columns: [
            { prop: INDEX_PROP, label: "我的序号" },
            { prop: "name", label: "名称" },
          ],
        }),
      ),
    ).toThrow(
      new RegExp(
        `column prop "${INDEX_PROP}" is reserved for the index column`,
      ),
    );
  });
});

describe("indexColumnStart", () => {
  it("resolves start from the object form, defaults to 1 otherwise", () => {
    expect(indexColumnStart(baseSheet({ indexColumn: { start: 100 } }))).toBe(
      100,
    );
    expect(indexColumnStart(baseSheet({ indexColumn: true }))).toBe(1);
    expect(indexColumnStart(baseSheet())).toBe(1);
  });
});
