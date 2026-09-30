// @vitest-environment happy-dom
import { beforeEach, describe, expect, it } from "vitest";
import { SheetRenderer } from "../render/renderer";
import {
  buildLayout,
  charsToPx,
  columnIndexLabel,
  computeVisibleRange,
  ptToPx,
} from "../render/layout";
import { compileStylesheet } from "../render/css";
import type { PreviewStyles } from "../types";
import { bigWorkbookBytes, sampleWorkbookBytes } from "./fixtures";
import { buildPreviewWorkbook } from "../parse/model";
import { readBuffer } from "@marcusok/xlsx-core";

async function parse(bytes: Uint8Array) {
  return buildPreviewWorkbook(await readBuffer(bytes), bytes);
}

describe("layout 几何", () => {
  it("单位换算", () => {
    expect(charsToPx(8.43)).toBe(64);
    expect(ptToPx(15)).toBe(20);
    expect(columnIndexLabel(0)).toBe("A");
    expect(columnIndexLabel(25)).toBe("Z");
    expect(columnIndexLabel(26)).toBe("AA");
    expect(columnIndexLabel(701)).toBe("ZZ");
  });

  it("buildLayout：列宽/行高/合并映射", async () => {
    const model = await parse(sampleWorkbookBytes());
    const l = buildLayout(model.sheets[0]);
    // 自定义列宽 20 → 145px；默认 8.43 → 64px；隐藏列 → 0
    expect(l.colWidths[0]).toBe(charsToPx(20));
    expect(l.colWidths[1]).toBe(0);
    expect(l.colWidths[2]).toBe(charsToPx(8.43));
    // 自定义行高 30pt；默认 15pt；隐藏行 0
    expect(l.rowHeights[3]).toBe(ptToPx(30));
    expect(l.rowHeights[1]).toBe(0);
    expect(l.rowHeights[0]).toBe(ptToPx(15));
    // 合并：anchor (0,0) 1×2；covered (0,1) → anchor
    expect(l.mergeByAnchor.get("0:0")).toEqual({
      row: 0,
      col: 0,
      rowSpan: 1,
      colSpan: 2,
    });
    expect(l.coveredBy.get("0:1")).toBe("0:0");
    expect(l.frozenRows).toBe(2);
    expect(l.frozenCols).toBe(1);
  });

  it("computeVisibleRange：合并召回（主格在窗口外）", async () => {
    // 大跨度合并 A1:A500：窗口在第 100 行附近时，合并主格必须被召回
    const bytes = bigWorkbookBytes(600, 6);
    const model = await parse(bytes);
    model.sheets[0].merges = [{ row: 0, col: 0, rowSpan: 500, colSpan: 1 }];
    const l = buildLayout(model.sheets[0]);
    const range = computeVisibleRange(l, l.rowTop[100], 0, 500, 300, 0);
    expect(range.rowStart).toBe(0); // 召回到合并主格
    expect(range.rowEnd).toBeGreaterThanOrEqual(101);
  });
});

describe("样式编译", () => {
  it("xf → CSS class（字体/填充/边框/对齐）", async () => {
    const model = await parse(sampleWorkbookBytes());
    const css = compileStylesheet(model.sheets[0].styles);
    expect(css).toContain(".xpv-xf-1{"); // theme 字体色（theme1 tint -0.25 → 黑，L=0 公式不可再暗）
    expect(css).toContain("color:#000000");
    expect(css).toContain("#99ccff"); // indexed 填充
    expect(css).toContain("border-bottom:1px solid #4472c4"); // 主题色边框
    // 对齐必须用 flex 属性：.xpv-cell 是 display:flex，text-align/vertical-align
    // 对匿名 flex item 无效（xf4 = horizontal center + vertical center）
    expect(css).toContain("justify-content:center");
    expect(css).toContain("align-items:center");
    expect(css).toContain("font-weight:700"); // 加粗
  });

  it("下划线+删除线合成一条 text-decoration，且不冲掉粗体/斜体", () => {
    // 回归守卫：此前合并两条 text-decoration 时用了整串赋值，把已累加的
    // font-weight/font-style 一并冲掉（bold+下划线+删除线丢失粗体）
    const styles: PreviewStyles = {
      fonts: [
        {
          name: null,
          size: null,
          bold: true,
          italic: true,
          underline: true,
          strike: true,
          color: null,
        },
      ],
      fills: [],
      borders: [],
      xfs: [
        {
          fontId: 0,
          fillId: 0,
          borderId: 0,
          numFmtCode: "General",
          alignment: null,
        },
      ],
    };
    const css = compileStylesheet(styles);
    expect(css).toContain("font-weight:700");
    expect(css).toContain("font-style:italic");
    expect(css).toContain("text-decoration:underline line-through");
    // 不再出现会按源序互相覆盖的两条独立声明（分号结尾才算是完整声明，
    // 合成形态 "underline line-through;" 不匹配这两个断言）
    expect(css).not.toContain("text-decoration:underline;");
    expect(css).not.toContain("text-decoration:line-through;");
  });

  it("渐变填充角度：xlsx degree 0（左→右）→ CSS 90deg", () => {
    const gradient: PreviewStyles = {
      fonts: [],
      borders: [],
      fills: [
        {
          kind: "gradient",
          degree: 0,
          stops: [
            { position: 0, color: "#ffffff" },
            { position: 1, color: "#000000" },
          ],
        },
      ],
      xfs: [
        {
          fontId: 0,
          fillId: 0,
          borderId: 0,
          numFmtCode: "General",
          alignment: null,
        },
      ],
    };
    expect(compileStylesheet(gradient)).toContain("linear-gradient(90deg");
    // degree 90（上→下）→ CSS 180deg
    const first = gradient.fills[0];
    gradient.fills[0] = {
      kind: "gradient",
      degree: 90,
      stops: first.kind === "gradient" ? first.stops : [],
    };
    expect(compileStylesheet(gradient)).toContain("linear-gradient(180deg");
  });

  it("对角线边框：方向随 diagonalUp，且与渐变填充分层不互相覆盖", () => {
    // 回归守卫 1：方向——diagonalUp = "/"（to top right），缺省/Down = "\"
    //（to bottom right）；回归守卫 2：渐变填充与对角线同落 background-image，
    // 此前分别写 background 简写与独立 background-image，后者按源序冲掉前者
    const styles: PreviewStyles = {
      fonts: [],
      fills: [
        {
          kind: "gradient",
          degree: 0,
          stops: [
            { position: 0, color: "#ffffff" },
            { position: 1, color: "#000000" },
          ],
        },
        { kind: "none" },
      ],
      borders: [
        {
          left: null,
          right: null,
          top: null,
          bottom: null,
          diagonal: { style: "thin", color: "#ff0000" },
          diagonalUp: true,
        },
        {
          left: null,
          right: null,
          top: null,
          bottom: null,
          // 无 diagonalUp 字段 = down（"\"，Excel 常见形态）
          diagonal: { style: "medium", color: "#000000" },
        },
      ],
      xfs: [
        {
          fontId: 0,
          fillId: 0,
          borderId: 0,
          numFmtCode: "General",
          alignment: null,
        },
        {
          fontId: 0,
          fillId: 1,
          borderId: 1,
          numFmtCode: "General",
          alignment: null,
        },
      ],
    };
    const css = compileStylesheet(styles);
    // xf0（渐变 + diagonalUp 对角线）：一条多层 background-image，对角线
    // 在上层、渐变（90deg）保留下层——两层共存于同一声明
    expect(css).toMatch(
      /\.xpv-xf-0\{[^}]*background-image:linear-gradient\(to top right[^;]*,\s*linear-gradient\(90deg/,
    );
    // xf1（无渐变、无 diagonalUp）：down 方向的独立声明
    expect(css).toMatch(
      /\.xpv-xf-1\{[^}]*background-image:linear-gradient\(to bottom right/,
    );
  });
});

describe("SheetRenderer", () => {
  let container: HTMLElement;

  beforeEach(() => {
    document.body.textContent = "";
    container = document.createElement("div");
    document.body.appendChild(container);
  });

  it("挂载即渲染首屏：格子/表头/页签", async () => {
    const model = await parse(sampleWorkbookBytes());
    const r = new SheetRenderer(container);
    r.render(model, {});
    // 首屏（回退视口 800×600）：主题字体格子等存在
    const cells = container.querySelectorAll(".xpv-cell");
    expect(cells.length).toBeGreaterThan(5);
    // 表头：列标 + 行号
    const headers = container.querySelectorAll(".xpv-hcell");
    const texts = [...headers].map((h) => h.textContent);
    expect(texts).toContain("A");
    expect(texts).toContain("1");
    // 页签：隐藏 sheet 不出
    const tabs = [...container.querySelectorAll(".xpv-tab")].map(
      (t) => t.textContent,
    );
    expect(tabs).toEqual(["S1", "S2"]);
    // 格式化文本进入 DOM（numFmt 165：¥1,234.50；numFmt 164：45678 → 2025-01-21）
    const texts2 = [...cells].map((c) => c.textContent);
    expect(texts2).toContain("¥1,234.50");
    expect(texts2).toContain("2025-01-21");
    // 合并主格 A1:B1 在冻结角块层（.xpv-layer 序：cells/frozenCol/frozenRow/
    // frozenCorner），跨列宽 = 145px + 隐藏列 0px
    const layers = container.querySelectorAll(".xpv-layer");
    const cornerCells = layers[3].querySelectorAll(".xpv-cell");
    const merged = [...cornerCells].find(
      (c) => c.textContent === "themeFont",
    ) as HTMLElement;
    expect(merged).toBeTruthy();
    expect(merged.style.width).toBe(`${charsToPx(20)}px`);
    r.destroy();
  });

  it("初始 sheet 越界数字：回退 activeSheetIndex，不再静默空白", async () => {
    const model = await parse(sampleWorkbookBytes());
    const r = new SheetRenderer(container);
    r.render(model, { sheet: 99 });
    // 此前 setSheet 提前 return、layout 为 null：预览静默空白且
    // onParsed/onError 均不触发
    expect(r.activeSheetIndex).toBe(0);
    expect(container.querySelectorAll(".xpv-cell").length).toBeGreaterThan(5);
    r.destroy();
  });

  it("切 sheet 与销毁", async () => {
    const model = await parse(sampleWorkbookBytes());
    const r = new SheetRenderer(container);
    r.render(model, {});
    expect(r.setSheet("S2")).toBe(1);
    expect(
      container.querySelectorAll(".xpv-tab.xpv-active")[0].textContent,
    ).toBe("S2");
    expect(r.setSheet("不存在")).toBe(-1);
    r.destroy();
    expect(
      document.body.contains(container.querySelector(".xpv-root") as Node),
    ).toBe(false);
    expect(document.querySelectorAll("style").length).toBe(0);
  });

  it("虚拟滚动窗口收敛（大文件不全量渲染）", async () => {
    const bytes = bigWorkbookBytes(5000, 8);
    const model = await parse(bytes);
    const r = new SheetRenderer(container);
    r.render(model, {});
    const cells = container.querySelectorAll(".xpv-cell").length;
    // 5000×8=40000 格，视口 ~800×600 → 数百格量级，绝不能全量
    expect(cells).toBeLessThan(2000);
    expect(cells).toBeGreaterThan(50);
    r.destroy();
  });

  it("冻结窗格分层存在", async () => {
    const model = await parse(sampleWorkbookBytes());
    const r = new SheetRenderer(container);
    r.render(model, {});
    // 冻结 2 行 1 列：冻结层有内容
    const frozenRowCells = container
      .querySelectorAll(".xpv-layer")[1]
      .querySelectorAll(".xpv-cell");
    expect(frozenRowCells.length).toBeGreaterThan(0);
    r.destroy();
  });

  it("冻结列标/行号在静态层，非冻结区在平移带（冻结表头 sticky）", async () => {
    const model = await parse(sampleWorkbookBytes());
    const r = new SheetRenderer(container);
    r.render(model, {});
    // 层序：cells/frozenCol/frozenRow/frozenCorner/headerRow/headerCol/
    // headerCorner/colHeadFrozen/rowHeadFrozen
    const layers = container.querySelectorAll(".xpv-layer");
    const colHeadFrozen = layers[7];
    const rowHeadFrozen = layers[8];
    // S1 冻结 1 列 2 行：冻结列标 "A" 在静态层，平移带（headerRow）从 "B" 起
    expect(
      [...colHeadFrozen.querySelectorAll(".xpv-hcell")].map(
        (h) => h.textContent,
      ),
    ).toEqual(["A"]);
    const headerRowLabels = [...layers[4].querySelectorAll(".xpv-hcell")].map(
      (h) => h.textContent,
    );
    expect(headerRowLabels).not.toContain("A");
    expect(headerRowLabels).toContain("B");
    // 冻结行号 1/2 在静态层，平移带（headerCol）从 3 起
    expect(
      [...rowHeadFrozen.querySelectorAll(".xpv-hcell")].map(
        (h) => h.textContent,
      ),
    ).toEqual(["1", "2"]);
    const headerColLabels = [...layers[5].querySelectorAll(".xpv-hcell")].map(
      (h) => h.textContent,
    );
    expect(headerColLabels).not.toContain("1");
    expect(headerColLabels).not.toContain("2");
    r.destroy();
  });

  it("覆盖层 transform 抵消滚动：钉住的轴 +scroll、跟随轴不动", async () => {
    const model = await parse(sampleWorkbookBytes());
    const r = new SheetRenderer(container);
    r.render(model, {});
    // 层序：cells/frozenCol/frozenRow/frozenCorner/headerRow/headerCol/
    // headerCorner/colHeadFrozen/rowHeadFrozen。覆盖层都随滚动位移（它们是
    // 滚动容器的绝对定位子元素），因此 transform 只抵消"要钉住"的轴：
    // 此前按"子元素不随滚动"写成了 -scroll/错轴，滚动时表头与冻结条带以
    // 两倍速度滑走（真实浏览器实测；happy-dom 只断言层归属时测不出）。
    const viewport = container.querySelector(".xpv-viewport") as HTMLElement;
    viewport.scrollTop = 160;
    viewport.scrollLeft = 120;
    viewport.dispatchEvent(new Event("scroll"));
    await new Promise((res) => requestAnimationFrame(() => res(null)));
    const layers = container.querySelectorAll<HTMLElement>(".xpv-layer");
    expect(layers[1].style.transform).toBe("translateX(120px)"); // 冻结列钉 X
    expect(layers[2].style.transform).toBe("translateY(160px)"); // 冻结行钉 Y
    expect(layers[4].style.transform).toBe("translateY(160px)"); // 列标钉 Y
    expect(layers[5].style.transform).toBe("translateX(120px)"); // 行号钉 X
    for (const i of [3, 6, 7, 8]) {
      // 冻结角块 / 左上角 / 冻结列标 / 冻结行号：双向钉
      expect(layers[i].style.transform).toBe("translate(120px, 160px)");
    }
    r.destroy();
  });

  it("网格线规则前置于 xf 类（同特异性级联，数据边框覆盖网格线）", async () => {
    const model = await parse(sampleWorkbookBytes());
    const r = new SheetRenderer(container);
    // S1 文件声明 showGridLines=0，显式开启以覆盖级联路径
    r.render(model, { showGridLines: true });
    // sheet 级样式元素 = 含 .xpv-xf- 类的那个（BASE_CSS 元素不含）
    const sheetCss = [...document.querySelectorAll("style")]
      .map((s) => s.textContent ?? "")
      .find((t) => t.includes(".xpv-xf-"));
    expect(sheetCss).toBeTruthy();
    const gridIdx = sheetCss!.indexOf(".xpv-cell{border-right");
    const xfIdx = sheetCss!.indexOf(".xpv-xf-");
    expect(gridIdx).toBeGreaterThan(-1);
    // 后置即回归：网格线会按源序覆盖 .xpv-xf-N 的右/下数据边框
    expect(gridIdx).toBeLessThan(xfIdx);
    r.destroy();
  });

  it("滚动不销毁冻结象限元素（窗口差量不变量）", async () => {
    const model = await parse(sampleWorkbookBytes());
    const r = new SheetRenderer(container);
    r.render(model, {});
    const layers = container.querySelectorAll(".xpv-layer");
    const corner = layers[3]; // cells/frozenCol/frozenRow/frozenCorner
    const el = [...corner.querySelectorAll(".xpv-cell")].find(
      (c) => c.textContent === "themeFont",
    ) as HTMLElement;
    expect(el).toBeTruthy();
    // 对照锚点：主区（冻结 2 行之外）的格子，滚出窗口后应被差量清理——
    // 该断言成立即证明 updateWindow 真实执行过，下面的"冻结元素未重建"
    // 因此不是空转
    const mainEl = layers[0].querySelectorAll(".xpv-cell")[0] as HTMLElement;
    expect(mainEl).toBeTruthy();
    const viewport = container.querySelector(".xpv-viewport") as HTMLElement;
    viewport.scrollTop = 600;
    viewport.dispatchEvent(new Event("scroll"));
    await new Promise((res) => requestAnimationFrame(() => res(null)));
    await new Promise((res) => requestAnimationFrame(() => res(null)));
    expect(mainEl.isConnected).toBe(false);
    // 同一元素引用仍连接：未被删除重建（此前缺陷：主象限清理以空前缀
    // 恒匹配所有键，冻结象限每滚动帧被误删后全量重建）
    expect(el.isConnected).toBe(true);
    expect(corner.contains(el)).toBe(true);
    r.destroy();
  });
});
