// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";
import { SheetRenderer } from "../render/renderer";
import type { PreviewWorkbook } from "../types";

/**
 * 容器可能来自 iframe 的 Document（与全局 document 不是同一个文档）。
 * 此时样式必须注入 container.ownerDocument.head，节点也必须用该 Document
 * 创建；否则 DOM 在 iframe 里、CSS 在父页面里，overflow/定位全部失效。
 */
describe("iframe-like Document scope", () => {
  const model: PreviewWorkbook = {
    activeSheetIndex: 0,
    dateSystem: "date1900",
    sheets: [
      {
        name: "S",
        visible: true,
        showGridLines: true,
        rightToLeft: false,
        rowCount: 1,
        colCount: 1,
        rows: [
          {
            index: 1,
            height: null,
            hidden: false,
            cells: [{ col: 0, type: "string", value: "x", styleIndex: null }],
          },
        ],
        colSpans: [],
        merges: [],
        frozenRows: 0,
        frozenCols: 0,
        styles: { fonts: [], fills: [], borders: [], xfs: [] },
      },
    ],
  };

  it("injects styles and creates nodes in the container's Document", () => {
    const otherDoc = document.implementation.createHTMLDocument("iframe");
    const container = otherDoc.createElement("div");
    otherDoc.body.appendChild(container);

    const renderer = new SheetRenderer(container);
    renderer.render(model, {});

    expect(otherDoc.querySelectorAll("style").length).toBeGreaterThan(0);
    expect(document.querySelectorAll("style").length).toBe(0);
    expect(otherDoc.querySelector(".xpv-root")).not.toBeNull();
    expect(otherDoc.querySelector(".xpv-cell")?.textContent).toBe("x");

    renderer.destroy();
    expect(otherDoc.querySelector(".xpv-root")).toBeNull();
  });
});
