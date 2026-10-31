import { describe, expect, it } from "vitest";
import {
  buildStylesOverlay,
  parseThemePalette,
  resolveColorSpec,
} from "../parse/styles-overlay";
import { applyTint, argbToCss, indexedColor } from "../color";
import {
  OFFICE_THEME,
  RICH_STYLES,
  bigWorkbookBytes,
  buildXlsx,
  colName,
  date1904Bytes,
  encryptedBytes,
  entityStylesBytes,
  fakeOle2Bytes,
  quotedFmtStylesBytes,
  richTextBytes,
  sampleWorkbookBytes,
  selfClosedStylesBytes,
  readBuffer,
} from "./fixtures";
import { buildPreviewWorkbook } from "../parse/model";
import { normalizeEngineError } from "../parse/engine-errors";
import {
  normalizeParseError,
  parseOnMainThread,
} from "../worker/worker-client";
import { csvToWorkbook, parseCsvText } from "../parse/csv";
import { buildLayout } from "../render/layout";
import { strToU8, zipSync } from "fflate";

// 畸形文件可以只声明一个超大行号/列跨度/合并，就让 layout 在主线程按声明
// 维度分配数 GB TypedArray（worker 超时覆盖不到布局）。模型层必须按 Excel
// 硬上限拒绝，且负/非有限几何要归一化，不能让负前缀和破坏窗口计算。
describe("Excel dimension limits (malformed files)", () => {
  async function expectCorrupt(bytes: Uint8Array, re: RegExp): Promise<void> {
    const wb = await readBuffer(bytes);
    let err: unknown;
    try {
      buildPreviewWorkbook(wb, bytes);
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(Error);
    expect((err as { code?: string }).code).toBe("CORRUPT");
    expect((err as Error).message).toMatch(re);
  }

  it("rejects a row index beyond Excel's 1,048,576-row limit", async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: "Big",
          xml: '<row r="100000000"><c r="A100000000" t="inlineStr"><is><t>x</t></is></c></row>',
        },
      ],
      styles: RICH_STYLES,
    });
    await expectCorrupt(bytes, /row 100000000/);
  });

  it("rejects a cell beyond column XFD (16,384)", async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: "Big",
          xml: '<row r="1"><c r="ZZZZZZ1" t="inlineStr"><is><t>x</t></is></c></row>',
        },
      ],
      styles: RICH_STYLES,
    });
    await expectCorrupt(bytes, /outside Excel's 1\.\.16384 range/);
  });

  it("rejects a merged range extending beyond the row limit", async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: "Big",
          xml: '<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>',
          extraAfter:
            '<mergeCells count="1"><mergeCell ref="A1:A100000000"/></mergeCells>',
        },
      ],
      styles: RICH_STYLES,
    });
    await expectCorrupt(bytes, /merged range outside Excel's/);
  });

  it("rejects a <col max> span beyond 16,384", async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: "Big",
          extraBefore: '<cols><col min="1" max="100000000" width="10"/></cols>',
          xml: '<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>',
        },
      ],
      styles: RICH_STYLES,
    });
    await expectCorrupt(bytes, /column span 1\.\.100000000/);
  });

  it("normalizes negative widths/heights instead of producing negative geometry", async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: "Neg",
          extraBefore: '<cols><col min="1" max="1" width="-100"/></cols>',
          xml: '<row r="1" ht="-50" customHeight="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>',
        },
      ],
      styles: RICH_STYLES,
    });
    const wb = await readBuffer(bytes);
    const model = buildPreviewWorkbook(wb, bytes);
    const sheet = model.sheets[0];
    expect(sheet.colSpans[0].width).toBe(8.43);
    expect(sheet.rows[0].height).toBeNull();
    const layout = buildLayout(sheet);
    expect(layout.totalWidth).toBeGreaterThan(0);
    expect(layout.totalHeight).toBeGreaterThan(0);
    expect(layout.colWidths[0]).toBeGreaterThan(0);
    expect(layout.rowHeights[0]).toBeGreaterThan(0);
  });
});

describe("color 基础", () => {
  it("ARGB → CSS（FF 前缀剥离、非 FF 转 rgba、非法 null）", () => {
    expect(argbToCss("FF0563C1")).toBe("#0563c1");
    expect(argbToCss("0563C1")).toBe("#0563c1");
    expect(argbToCss("80054321")).toBe("rgba(5, 67, 33, 0.5)");
    expect(argbToCss("XYZ")).toBeNull();
    expect(argbToCss(null)).toBeNull();
  });
  it("tint 变亮/变暗单调且边界安全", () => {
    const base = "4472C4";
    const lighter = applyTint(base, 0.6);
    const darker = applyTint(base, -0.25);
    // HSL 亮度单调：lighter 更亮（RGB 分量和更大），darker 更暗
    const sum = (h: string) =>
      [0, 2, 4].reduce((s, i) => s + parseInt(h.slice(i, i + 2), 16), 0);
    expect(sum(lighter)).toBeGreaterThan(sum(base));
    expect(sum(darker)).toBeLessThan(sum(base));
    expect(applyTint(base, 0)).toBe(base.toLowerCase());
    expect(applyTint(base, 1)).toMatch(/^(ff){3}$/); // 全白
    expect(applyTint(base, -1)).toBe("000000"); // 全黑
  });
  it("indexed 调色板", () => {
    expect(indexedColor(44)).toBe("99ccff");
    // 51/62/63：此前 51（ccff00）与 52（ffcc00）缺失导致 53-62 整体前移
    // 两位、62/63 值错（对照 ECMA-376 §18.8.27 附表修正的回归锚点）
    expect(indexedColor(51)).toBe("ccff00");
    expect(indexedColor(52)).toBe("ffcc00");
    expect(indexedColor(53)).toBe("ff9900");
    expect(indexedColor(62)).toBe("993366");
    expect(indexedColor(63)).toBe("333399");
    expect(indexedColor(64)).toBe("000000");
    expect(indexedColor(999)).toBeNull();
  });
});

describe("主题色覆盖层（读取侧 null 的找回）", () => {
  const bytes = sampleWorkbookBytes();

  it("styles.xml 颜色规格按文档序解出", () => {
    const { overlay } = buildStylesOverlay(bytes);
    expect(overlay).not.toBeNull();
    // font[1] = theme1 tint -0.25
    expect(overlay!.fonts[1].color).toEqual({
      kind: "theme",
      theme: 1,
      tint: -0.25,
    });
    // fill[2] = theme4 tint 0.6（accent1）
    expect(overlay!.fills[2].fgColor).toEqual({
      kind: "theme",
      theme: 4,
      tint: 0.6,
    });
    // fill[3] = indexed 44
    expect(overlay!.fills[3].fgColor).toEqual({
      kind: "indexed",
      indexed: 44,
    });
    // border#0 自闭合（真实形态）解出空项；border[1].bottom = theme4
    expect(overlay!.borders[0]).toEqual({
      left: null,
      right: null,
      top: null,
      bottom: null,
      diagonal: null,
      diagonalUp: false,
      diagonalDown: false,
    });
    expect(overlay!.borders[1].bottom?.color).toEqual({
      kind: "theme",
      theme: 4,
    });
  });

  it("命名空间前缀 + 单引号属性的 styles.xml 仍能完整解析", () => {
    // 合法 XML 的两种常见变体：元素带前缀（<x:fonts>）与单引号属性
    // （patternType='solid'）。旧解析器会产出非 null 的空覆盖层，使模型层
    // 不再回退引擎数组，样式整片静默丢失。
    const styles = `<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:fonts count="2"><x:font><x:b/><x:sz val='12'/><x:name val='Arial'/><x:color rgb='FFFF0000'/></x:font><x:font/></x:fonts><x:fills count="1"><x:fill><x:patternFill patternType='solid'><x:fgColor rgb='FF00FF00'/></x:patternFill></x:fill></x:fills><x:borders count="1"><x:border><x:left style='thin'><x:color rgb='FF0000FF'/></x:left></x:border></x:borders><x:cellXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></x:cellXfs></x:styleSheet>`;
    const bytes2 = buildXlsx({
      sheets: [
        {
          name: "P",
          xml: '<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>',
        },
      ],
      styles,
    });
    const { overlay } = buildStylesOverlay(bytes2);
    expect(overlay).not.toBeNull();
    expect(overlay!.fonts).toHaveLength(2);
    expect(overlay!.fonts[0]).toMatchObject({
      bold: true,
      size: 12,
      name: "Arial",
      color: { kind: "rgb", rgb: "FFFF0000" },
    });
    expect(overlay!.fills).toHaveLength(1);
    expect(overlay!.fills[0].patternType).toBe("solid");
    expect(overlay!.fills[0].fgColor).toEqual({
      kind: "rgb",
      rgb: "FF00FF00",
    });
    expect(overlay!.borders).toHaveLength(1);
    expect(overlay!.borders[0].left?.style).toBe("thin");
    expect(overlay!.borders[0].left?.color).toEqual({
      kind: "rgb",
      rgb: "FF0000FF",
    });
  });

  it("解析结果与 section count 不一致时降级（返回 null，交引擎数组）", () => {
    // 覆盖层数组必须与 cellXfs 的 id 严格 1:1；数量对不上时宁可整体
    // 放弃，也不能交出错位的样式表。
    const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs></styleSheet>`;
    const bytes2 = buildXlsx({
      sheets: [
        {
          name: "M",
          xml: '<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>',
        },
      ],
      styles,
    });
    expect(buildStylesOverlay(bytes2).overlay).toBeNull();
  });

  it("theme1.xml 调色板解出", () => {
    const palette = parseThemePalette(`<?xml version="1.0"?>${OFFICE_THEME}`);
    expect(palette.accent1).toBe("4472C4");
    expect(palette.dk1).toBe("000000");
    expect(palette.folHlink).toBe("954F72");
  });

  it("规格解析成最终 CSS 色", () => {
    const palette = parseThemePalette(`<?xml version="1.0"?>${OFFICE_THEME}`);
    // 黑色（theme 1 = dk1）负 tint 按公式仍为黑（L=0 不可再暗）
    expect(
      resolveColorSpec({ kind: "theme", theme: 1, tint: -0.25 }, palette),
    ).toBe("#000000");
    expect(resolveColorSpec({ kind: "theme", theme: 4 }, palette)).toBe(
      "#4472c4",
    );
    expect(resolveColorSpec({ kind: "indexed", indexed: 44 }, palette)).toBe(
      "#99ccff",
    );
    expect(resolveColorSpec({ kind: "rgb", rgb: "FFFF0000" }, palette)).toBe(
      "#ff0000",
    );
    expect(resolveColorSpec(null, palette)).toBeNull();
  });

  it("单独解析 styles.xml 字符串（无 zip 场景）不抛错", () => {
    expect(() => buildStylesOverlay(strToU8("not a zip"))).not.toThrow();
  });
});

describe("下划线 val 枚举与对角线标志（ECMA-376 语义）", () => {
  // <u> 的 val 是 ST_UnderlineValues 枚举（缺省 "single"），不是布尔：
  // 显式 val="single" 与 <u/> 等价（LibreOffice 等第三方导出器常见形态），
  // 此前按 flagElem 布尔口径解析会把这些文件的下划线整体丢掉。
  // 对角线同理：ECMA-376 要求 <border> 带 diagonalUp/Down 标志才显示，
  // 仅有 <diagonal style> 不画线（此前凭空渲染且方向固定 "/"）。
  const STYLES = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><u val="single"/><sz val="11"/><name val="Calibri"/></font><font><u val="none"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="3"><border><left/><right/><top/><bottom/><diagonal style="thin"/></border><border diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"/></border><border diagonalUp="1"><left/><right/><top/><bottom/><diagonal style="thin"/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;
  const bytes = buildXlsx({
    sheets: [
      {
        name: "U",
        xml: `<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>`,
      },
    ],
    styles: STYLES,
  });

  it('<u val="single"/> 等价于 <u/>：有下划线；val="none" 显式无', () => {
    const { overlay } = buildStylesOverlay(bytes);
    expect(overlay!.fonts[0].underline).toBe(false); // 无 <u> 元素
    expect(overlay!.fonts[1].underline).toBe(true); // val="single"（缺省值的显式形态）
    expect(overlay!.fonts[2].underline).toBe(false); // val="none"
  });

  it('无 Up/Down 标志的对角线样式不渲染；diagonalDown → "\\"、diagonalUp → "/"', async () => {
    const { overlay } = buildStylesOverlay(bytes);
    expect(overlay!.borders[0]).toMatchObject({
      diagonal: { style: "thin" },
      diagonalUp: false,
      diagonalDown: false,
    });
    expect(overlay!.borders[1]).toMatchObject({
      diagonalDown: true,
      diagonalUp: false,
    });
    expect(overlay!.borders[2]).toMatchObject({
      diagonalUp: true,
      diagonalDown: false,
    });
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    const st = model.sheets[0].styles;
    // 模型层：无标志 → diagonal 置 null（Excel 语义：不画 Excel 不画的线）
    expect(st.borders[0].diagonal).toBeNull();
    // diagonalDown → 保留样式、方向为 "\"（diagonalUp=false）
    expect(st.borders[1].diagonal).toEqual({ style: "thin", color: null });
    expect(st.borders[1].diagonalUp).toBe(false);
    // diagonalUp → 方向 "/"（diagonalUp=true）
    expect(st.borders[2].diagonalUp).toBe(true);
  });
});

describe("模型构建（含覆盖层合并）", () => {
  it("主题色/indexed 颜色在模型里已是解析后的 RGB", async () => {
    const { readBuffer } = await import("@marcusok/xlsx-core");
    const wb = await readBuffer(sampleWorkbookBytes());
    const model = buildPreviewWorkbook(wb, sampleWorkbookBytes());
    const styles = model.sheets[0].styles;
    // font[1].color：theme1（黑）tint -0.25 → 黑（L=0 不可再暗，公式忠实）
    expect(styles.fonts[1].color).toBe("#000000");
    // fill[2]：theme4 tint 0.6 → accent1 变亮
    expect(styles.fills[2]).toMatchObject({
      kind: "solid",
      fgColor: "#b4c7e7",
    });
    // fill[3]：indexed 44
    expect(styles.fills[3]).toMatchObject({
      kind: "solid",
      fgColor: "#99ccff",
    });
    // border[1].bottom：theme4 → accent1
    expect(styles.borders[1].bottom).toEqual({
      style: "thin",
      color: "#4472c4",
    });
  });

  it("行列/合并/冻结/隐藏归一化", async () => {
    const { readBuffer } = await import("@marcusok/xlsx-core");
    const bytes = sampleWorkbookBytes();
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    const s = model.sheets[0];
    expect(s.name).toBe("S1");
    expect(s.merges).toEqual([{ row: 0, col: 0, rowSpan: 1, colSpan: 2 }]);
    expect(s.frozenRows).toBe(2);
    expect(s.frozenCols).toBe(1);
    expect(s.showGridLines).toBe(false);
    // 隐藏行：index 2，hidden true
    const row2 = s.rows.find((r) => r.index === 2);
    expect(row2?.hidden).toBe(true);
    // 隐藏列 span
    expect(s.colSpans.find((c) => c.min === 2)?.hidden).toBe(true);
    // 稀疏：行 2 之外 1-4 全在
    expect(s.rows.map((r) => r.index)).toEqual([1, 2, 3, 4]);
    // 行高
    expect(s.rows.find((r) => r.index === 4)?.height).toBe(30);
    expect(s.rows.find((r) => r.index === 1)?.height).toBeNull();
    // 单元格类型归一
    const row3 = s.rows.find((r) => r.index === 3)!.cells;
    expect(row3.find((c) => c.col === 0)).toMatchObject({
      type: "number",
      value: "3",
    });
    expect(row3.find((c) => c.col === 1)?.type).toBe("boolean");
    expect(row3.find((c) => c.col === 2)).toMatchObject({
      type: "error",
      value: "#DIV/0!",
    });
    // numFmt 解析进 xf：xf5 = ¥#,##0.00、xf4 = yyyy-mm-dd
    expect(s.styles.xfs[5].numFmtCode).toBe("¥#,##0.00");
    expect(s.styles.xfs[4].numFmtCode).toBe("yyyy-mm-dd");
    // 对齐
    expect(s.styles.xfs[4].alignment).toMatchObject({
      horizontal: "center",
      wrapText: true,
      indent: 1,
    });
  });

  it("空行自定义行高保留（引擎输出无 cell 但带 height 的行）", async () => {
    // 回归：<row ht customHeight> 无 cell 的行若被稀疏化跳过，渲染层会把
    // 该行回落默认行高（拉高留白的空行整体塌掉）
    const bytes = buildXlsx({
      sheets: [
        {
          name: "EH",
          xml: `<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row><row r="2" ht="42" customHeight="1"/><row r="3"><c r="A3" t="inlineStr"><is><t>b</t></is></c></row>`,
        },
      ],
      styles: RICH_STYLES,
    });
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    const row2 = model.sheets[0].rows.find((r) => r.index === 2);
    expect(row2).toMatchObject({ height: 42, hidden: false });
    expect(row2?.cells).toEqual([]);
  });

  it("numFmt formatCode 的 XML 实体解码（引擎读取侧原样返回 &quot;）", async () => {
    // 回归（2026-10 真实浏览器复现）：引号字面量格式在 styles.xml 属性里
    // 转义为 &quot;，引擎 readBuffer 不做实体解码。模型层不补解码时，
    // 渲染出 "2026&quot;1&quot;15&quot;" 一类乱码（exporter 自家写出的文件
    // 读回即触发）。断言覆盖模型层（numFmtCode 解码）与渲染全链路（两条
    // 格式化路径：日期段自实现 + 数字段边字面量回注）。
    const bytes = quotedFmtStylesBytes();
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    const { formatCellValue } = await import("../numfmt/format");
    const xfs = model.sheets[0].styles.xfs;
    // 模型层：formatCode 已解码为带真实引号的格式串
    expect(xfs[0].numFmtCode).toBe('yyyy"年"m"月"d"日"');
    expect(xfs[1].numFmtCode).toBe('0"元"');
    // 渲染层：日期段（自实现路径）按字面量渲染，无 &quot; 残留
    const date = formatCellValue(
      "number",
      "45678",
      xfs[0].numFmtCode,
      model.dateSystem,
    );
    expect(date.text).toMatch(/^\d{4}年\d{1,2}月\d{1,2}日$/);
    // 渲染层：数字段（边字面量回注路径）后缀字面量精确匹配
    const money = formatCellValue(
      "number",
      "123",
      xfs[1].numFmtCode,
      model.dateSystem,
    );
    expect(money.text).toBe("123元");
  });

  it("sheet 可见性与 activeTab", async () => {
    const { readBuffer } = await import("@marcusok/xlsx-core");
    const bytes = sampleWorkbookBytes();
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    expect(model.sheets.map((s) => [s.name, s.visible])).toEqual([
      ["S1", true],
      ["S2", true],
      ["Hidden", false],
    ]);
    expect(model.activeSheetIndex).toBe(0);
  });

  it("activeTab 指向隐藏 sheet：顺延到首个可见 sheet", async () => {
    // 第 3 个 sheet（index 2）是 hidden——此前 activeSheetIndex=2 会渲染
    // 隐藏内容，而页签栏按 visible 过滤后无对应页签可切走
    const bytes = sampleWorkbookBytes(2);
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    expect(model.activeSheetIndex).toBe(0);
  });

  it("全隐藏 workbook：activeTab 保持原值（无目标可回退）", async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: "A",
          state: "hidden",
          xml: `<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row>`,
        },
        {
          name: "B",
          state: "hidden",
          xml: `<row r="1"><c r="A1" t="inlineStr"><is><t>b</t></is></c></row>`,
        },
      ],
      styles: RICH_STYLES,
      activeTab: 1,
    });
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    expect(model.activeSheetIndex).toBe(1);
  });

  it("date1904 标记透传", async () => {
    const { readBuffer } = await import("@marcusok/xlsx-core");
    const bytes = date1904Bytes();
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    expect(model.dateSystem).toBe("date1904");
  });

  it("富文本（rows 路径 value 已是拼接文本）", async () => {
    const { readBuffer } = await import("@marcusok/xlsx-core");
    const bytes = richTextBytes();
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    expect(model.sheets[0].rows[0].cells[0].value).toBe("Hello");
    expect(model.sheets[0].rows[0].cells[1].value).toBe("plain");
  });

  it("合并区超出内容边界：rowCount/colCount 并入合并末行/列", async () => {
    // Excel 对纯合并不写覆盖格（引擎 rows 只回主格行）——此前按内容归约
    // 得 1×1，渲染层把 A1:C3 合并截断为 1×1 视觉尺寸
    const bytes = buildXlsx({
      sheets: [
        {
          name: "M",
          xml: `<row r="1"><c r="A1" t="inlineStr"><is><t>merged</t></is></c></row>`,
          extraAfter:
            '<mergeCells count="1"><mergeCell ref="A1:C3"/></mergeCells>',
        },
      ],
      styles: RICH_STYLES,
    });
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    const s = model.sheets[0];
    expect(s.rowCount).toBe(3);
    expect(s.colCount).toBe(3);
    expect(s.merges).toEqual([{ row: 0, col: 0, rowSpan: 3, colSpan: 3 }]);
  });

  it("大样本端到端（模型规模）", async () => {
    const { readBuffer } = await import("@marcusok/xlsx-core");
    const bytes = bigWorkbookBytes(2000, 10);
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    const s = model.sheets[0];
    expect(s.rowCount).toBe(2000);
    expect(s.colCount).toBe(10);
    expect(s.rows[1999].cells[9]).toMatchObject({ col: 9, type: "number" });
    expect(colName(10)).toBe("J");
  });
});

describe("自闭合样式元素（真实 Excel 产物形态的 id 对齐）", () => {
  // 引擎读取侧跳过 <border/>、<font/>、<fill/> 自闭合元素，其数组比文件
  // 各少一项、与 cellXfs 的 id 引用整体错位（实测：borderId=0 的默认单元
  // 格拿到文件 border#1 的样式、末位引用越界丢失）。模型经覆盖层按文档序
  // 完整重建（含空项）后 id 严格对齐——本组即该修复的回归锚点。
  it("fonts/fills/borders 含自闭合空项，长度与 cellXfs 的 id 语义一致", () => {
    const bytes = selfClosedStylesBytes();
    const { overlay } = buildStylesOverlay(bytes);
    expect(overlay!.fonts).toHaveLength(2);
    expect(overlay!.fills).toHaveLength(3);
    expect(overlay!.borders).toHaveLength(3);
    // 自闭合空项形态
    expect(overlay!.fonts[0]).toEqual({
      name: null,
      size: null,
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      color: null,
    });
    expect(overlay!.fills[0]).toEqual({
      patternType: null,
      fgColor: null,
      bgColor: null,
      gradient: null,
    });
    // 非空项：style 与颜色规格都在
    expect(overlay!.borders[1].bottom).toEqual({
      style: "thin",
      color: { kind: "theme", theme: 1 },
    });
  });

  it("模型三个集合与 cellXfs 引用严格对齐（修复前整体前移错位）", async () => {
    const bytes = selfClosedStylesBytes();
    const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
    const st = model.sheets[0].styles;
    expect(st.borders).toHaveLength(3);
    expect(st.fonts).toHaveLength(2);
    expect(st.fills).toHaveLength(3);
    // 集合本体：空项与显式项（theme1 → 黑）
    expect(st.borders[0].bottom).toBeNull();
    expect(st.borders[1].bottom).toEqual({ style: "thin", color: "#000000" });
    expect(st.borders[2].bottom).toEqual({ style: "medium", color: "#ff0000" });
    expect(st.fonts[1].bold).toBe(true);
    expect(st.fills[0]).toEqual({ kind: "none" });
    expect(st.fills[2]).toEqual({ kind: "solid", fgColor: "#ffff00" });
    // 单元格经 cellXfs 的 id 取到文件语义的样式：
    // A1(xf1)=thin 黑、B1(xf2)=medium 红、C1(xf3)=粗体、D1(xf4)=黄填充
    const cells = model.sheets[0].rows[0].cells;
    const xfOf = (i: number) => st.xfs[cells[i].styleIndex!];
    expect(st.borders[xfOf(0).borderId].bottom).toEqual({
      style: "thin",
      color: "#000000",
    });
    expect(st.borders[xfOf(1).borderId].bottom).toEqual({
      style: "medium",
      color: "#ff0000",
    });
    expect(st.fonts[xfOf(2).fontId].bold).toBe(true);
    expect(st.fills[xfOf(3).fillId]).toEqual({
      kind: "solid",
      fgColor: "#ffff00",
    });
  });
});

describe("实体解码健壮性（恶意字符引用不得逃出覆盖层）", () => {
  it("超范围十进制实体（>0x10FFFF）替换为 U+FFFD，端到端解析不失败", async () => {
    const bytes = entityStylesBytes("&#1114112;");
    // 覆盖层本身不抛（此前 RangeError 从这里逃出 try 只包 unzip 的保护圈）；
    // U+FFFD 不是数字，numVal 落回 null 兜底
    const { overlay } = buildStylesOverlay(bytes);
    expect(overlay?.fonts[0]?.size).toBeNull();
    // 端到端：整份文件照常解析成功，而不是 UNKNOWN 失败
    const r = await parseOnMainThread({ bytes });
    expect(r.ok).toBe(true);
  });

  it("代理区实体（孤立 surrogate）替换为 U+FFFD，解析不失败", async () => {
    const r = await parseOnMainThread({ bytes: entityStylesBytes("&#xD800;") });
    expect(r.ok).toBe(true);
  });

  it("十六进制实体（XML CharRef 的 &#x 形态）按规范解码", () => {
    // &#x31; = "1"：若不解码，Number("&#x31;") 为 NaN、size 应为 null
    const { overlay } = buildStylesOverlay(entityStylesBytes("&#x31;"));
    expect(overlay?.fonts[0]?.size).toBe(1);
  });

  it("合法十进制实体照常解码（回归确认）", () => {
    // 两个 &#50; 拼出 "22"：验证多实体连写与既有十进制路径
    const { overlay } = buildStylesOverlay(entityStylesBytes("&#50;&#50;"));
    expect(overlay?.fonts[0]?.size).toBe(22);
  });
});

describe("解析入口（主线程路径 + 错误归一）", () => {
  it("正常解析出模型", async () => {
    const r = await parseOnMainThread({ bytes: sampleWorkbookBytes() });
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.workbook.sheets[0].name).toBe("S1");
    }
  });

  it("CSV 文本分流", async () => {
    const r = await parseOnMainThread({ bytes: strToU8("a,b\n1,2\n3,4\n") });
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.workbook.sheets[0].rowCount).toBe(3);
    }
  });

  it("以 PK 开头的 CSV 不会被误判为 zip / legacy .xls", async () => {
    // "PK,id" 是很常见的主键表头；只匹配两字节 PK 会把文本当 ZIP，引擎
    // 报 "expected ZIP or OLE2 header"，随后被归一成 LEGACY_FORMAT。
    const r = await parseOnMainThread({
      bytes: strToU8("PK,id\n1,2\n3,4\n"),
    });
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.workbook.sheets[0].name).toBe("CSV");
      expect(r.workbook.sheets[0].rowCount).toBe(3);
    }
  });

  it("伪造 OLE2 → LEGACY_FORMAT（.xls 友好报错）", async () => {
    const r = await parseOnMainThread({ bytes: fakeOle2Bytes() });
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.code).toBe("LEGACY_FORMAT");
      expect(r.message).toMatch(/\.xls/i);
    }
  });

  it("非 xlsx 的 zip（.ods/.docx 同形）→ CORRUPT,而非 UNKNOWN", async () => {
    // 实测：引擎抛 MISSING_PART（zip 里没有 xl/workbook.xml）。此前只有 4 个
    // 引擎码被识别，这里会落到 UNKNOWN 并透出引擎原文
    const ods = zipSync({
      mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
      "content.xml": strToU8("<office:document-content/>"),
    });
    const r = await parseOnMainThread({ bytes: ods });
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.code).toBe("CORRUPT");
      expect(r.message).toMatch(/valid \.xlsx/i);
    }
  });

  it("截断/损坏的 zip（ZIP 头在、目录坏）→ CORRUPT,而非 UNKNOWN", async () => {
    // 实测：引擎抛 ZIP_READ（Could not find EOCD）——真实截断的 xlsx 同形
    const truncated = new Uint8Array([
      0x50,
      0x4b,
      0x03,
      0x04,
      ...new Uint8Array(200).fill(0x41),
    ]);
    const r = await parseOnMainThread({ bytes: truncated });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.code).toBe("CORRUPT");
  });

  it("零 sheet 的 zip → CORRUPT（不再静默产出空模型）", async () => {
    // 实测：zip 里只有 workbook.xml（含 XML 已损坏的形态）时引擎"成功"返回
    // 0 个 sheet；渲染层对空 sheet 列表无 sheet 可切 → 静默空白且 onParsed/
    // onError 均不触发。模型层守卫改为确定性域错误
    const noSheets = zipSync({
      "xl/workbook.xml": strToU8('<?xml version="1.0"?><workbook/>'),
    });
    const r = await parseOnMainThread({ bytes: noSheets });
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.code).toBe("CORRUPT");
      expect(r.message).toMatch(/no sheets/i);
    }
  });

  it("加密文件无密码 → PASSWORD_PROTECTED；带密码读出", async () => {
    const enc = await encryptedBytes("pw123");
    const noPw = await parseOnMainThread({ bytes: enc });
    expect(noPw.ok).toBe(false);
    if (!noPw.ok) expect(noPw.code).toBe("PASSWORD_PROTECTED");

    const withPw = await parseOnMainThread({ bytes: enc, password: "pw123" });
    expect(withPw.ok).toBe(true);
    if (withPw.ok) {
      expect(withPw.workbook.sheets[0].rows[0].cells[0].value).toBe("secret");
    }
  });

  it("normalizeParseError：非错误码异常 → UNKNOWN", () => {
    const r = normalizeParseError(new Error("boom"));
    expect(r).toMatchObject({ ok: false, code: "UNKNOWN", message: "boom" });
  });

  it("normalizeEngineError：UNRECOGNIZED_FORMAT 归一到公共错误码（worker/主线程双路径一致）", () => {
    const e = (code: string, message: string) =>
      Object.assign(new Error(message), { code });
    // 非 OLE2 的不可识别格式 → CORRUPT（此前 worker 路径原样透传
    // "UNRECOGNIZED_FORMAT"，不在 PreviewErrorCode 联合内）
    expect(
      normalizeEngineError(
        e("UNRECOGNIZED_FORMAT", "Unrecognized format: garbage"),
      ),
    ).toMatchObject({ code: "CORRUPT", domain: true });
    // OLE2 信息（真实/伪造 .xls）→ LEGACY_FORMAT
    expect(
      normalizeEngineError(
        e(
          "UNRECOGNIZED_FORMAT",
          "File too small for OLE2 header: got 80 bytes",
        ),
      ),
    ).toMatchObject({ code: "LEGACY_FORMAT", domain: true });
    expect(normalizeEngineError(e("LEGACY_FORMAT", "legacy"))).toMatchObject({
      code: "LEGACY_FORMAT",
      domain: true,
    });
    expect(
      normalizeEngineError(e("PASSWORD_PROTECTED", "protected")),
    ).toMatchObject({ code: "PASSWORD_PROTECTED", domain: true });
    // 环境性异常：不标 domain（允许回退主线程重试）
    expect(normalizeEngineError(new Error("wasm exploded"))).toMatchObject({
      code: "UNKNOWN",
      domain: false,
    });
  });

  it("normalizeEngineError：结构类码 → CORRUPT，WASM 类码 → WASM", () => {
    const e = (code: string, message: string) =>
      Object.assign(new Error(message), { code });
    // 结构类（实测：非 xlsx 的 zip 抛 MISSING_PART、截断包抛 ZIP_READ）：
    // 确定性的文件性问题，标 domain（主线程复跑必然同错）
    for (const code of ["ZIP_READ", "ZIP_ENTRY", "MISSING_PART", "XML_PARSE"]) {
      expect(normalizeEngineError(e(code, "raw engine text"))).toMatchObject({
        code: "CORRUPT",
        domain: true,
      });
    }
    // WASM 类：环境性失败，不标 domain（worker 内失败不代表主线程也失败）
    for (const code of ["WASM", "WASM_ERROR", "WASM_INIT_FAILED"]) {
      expect(normalizeEngineError(e(code, "raw engine text"))).toMatchObject({
        code: "WASM",
        domain: false,
      });
    }
    // 预览自产码透传（模型层抛的零 sheet CORRUPT 走这条路），不再折成 UNKNOWN
    expect(normalizeEngineError(e("CORRUPT", "no sheets"))).toMatchObject({
      code: "CORRUPT",
      domain: true,
    });
    expect(normalizeEngineError(e("UNSUPPORTED", "nope"))).toMatchObject({
      code: "UNSUPPORTED",
      domain: true,
    });
  });

  it("XML/HTML 伪 xls 不再当 CSV 渲染（UNSUPPORTED 友好报错）", async () => {
    const xml = await parseOnMainThread({
      bytes: strToU8(
        '<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"></Workbook>',
      ),
    });
    expect(xml.ok).toBe(false);
    if (!xml.ok) expect(xml.code).toBe("UNSUPPORTED");

    const html = await parseOnMainThread({
      bytes: strToU8(
        "<html><body><table><tr><td>a</td></tr></table></body></html>",
      ),
    });
    expect(html.ok).toBe(false);
    if (!html.ok) expect(html.code).toBe("UNSUPPORTED");
  });

  it("高 0x80+ 占比的二进制不再当 CSV 渲染（fatal 解码双失败 → UNSUPPORTED）", async () => {
    // PNG 魔数 + 0xFF 填充：所有 ≥0x80 字节都被可打印率宽容计数，比例虚高
    // 过 95% 阈值；但连续 0xFF 既不是 UTF-8 也不落在 GB18030 双/四字节的
    // 任何合法位置，fatal 解码必失败——归 unknown 而非渲染成乱码行。
    const binary = new Uint8Array([
      0x89,
      0x50,
      0x4e,
      0x47,
      0x0d,
      0x0a,
      0x1a,
      0x0a,
      ...new Array(512).fill(0xff),
    ]);
    const r = await parseOnMainThread({ bytes: binary });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.code).toBe("UNSUPPORTED");
  });

  it("GBK 编码 CSV 仍按文本解析（含 2KB 边界截断在双字节中间的场景）", async () => {
    // "姓名,数量" 的 GBK 编码字节（手工锁定，避免依赖 Node iconv）：
    // 0xD0 0xD5 0xC3 0xFB, 0xCA 0xFD 0xC1 0xBF
    const gbkName = [0xd0, 0xd5, 0xc3, 0xfb, 0x2c, 0xca, 0xfd, 0xc1, 0xbf];
    // 前 2047 字节 ASCII（换行铺满）+ GBK 双字节跨 2KB 采样边界：
    // 采样 n=2048 恰好吞下首字节 0xD0 而切掉次字节 → 截断兜底去尾重试。
    const prefix = "a,b\n".repeat(Math.ceil(2047 / 4)).slice(0, 2047);
    const gkbBytes = new Uint8Array([
      ...strToU8(prefix),
      ...gbkName,
      ...strToU8("\n1,2\n"),
    ]);
    const wb = await parseOnMainThread({ bytes: gkbBytes });
    expect(wb.ok).toBe(true);
    if (wb.ok) {
      // GBK 经 decodeBytes 回退链解码为 "姓名,数量" 行
      const row = wb.workbook.sheets[0].rows.find((r) => r.index > 512);
      expect(row).toBeTruthy();
      expect(wb.workbook.sheets[0].rowCount).toBeGreaterThan(512);
    }

    // 无边界的普通 GBK CSV（体量小于采样窗口）不受影响
    const small = await parseOnMainThread({
      bytes: new Uint8Array([...gbkName, ...strToU8("\n1,2")]),
    });
    expect(small.ok).toBe(true);
    if (small.ok) {
      expect(small.workbook.sheets[0].rows[0].cells[0].value).toBe("姓名");
    }
  });
});

describe("CSV 解析", () => {
  it("RFC 4180：引号转义与跨行字段", () => {
    const rows = parseCsvText('a,"b,c"\n"x""y",2\n', ",");
    expect(rows).toEqual([
      ["a", "b,c"],
      ['x"y', "2"],
    ]);
  });
  it("分隔符嗅探", () => {
    const wb = csvToWorkbook(strToU8("a;b\n1;2\n"));
    expect(wb.sheets[0].rows[0].cells.map((c) => c.value)).toEqual(["a", "b"]);
  });
  it("TSV 字段内的逗号不会让分隔符误选为逗号", () => {
    // 逗号总频次（3）高于制表符（2）；按频次投票会把整表按逗号拆成
    // 4 列。按行一致性评分时 tab 是唯一稳定的两列结构。
    const wb = csvToWorkbook(
      strToU8("desc\tamount\nWidget, large, red, blue\t10\nRow2\t20\n"),
    );
    const sheet = wb.sheets[0];
    expect(sheet.colCount).toBe(2);
    expect(sheet.rows[1].cells.map((c) => c.value)).toEqual([
      "Widget, large, red, blue",
      "10",
    ]);
  });
  it("BOM 剥离与数字类型推断", () => {
    const wb = csvToWorkbook(
      new Uint8Array([0xef, 0xbb, 0xbf, 0x31, 0x2c, 0x32, 0x0a]),
    );
    const cells = wb.sheets[0].rows[0].cells;
    expect(cells[0]).toMatchObject({ type: "number", value: "1" });
    expect(cells[1]).toMatchObject({ type: "number", value: "2" });
  });
  it("尾点数字按数字推断（Excel 口径，123. 即 123）", () => {
    // Excel 输入栏里 "123." 视为数字 123；旧正则（\.\d+ 要求点后有数字）
    // 把它留在文本类，数字列对齐/格式化全部失效
    const wb = csvToWorkbook(strToU8("123.,4.5e2,6.\n"));
    const cells = wb.sheets[0].rows[0].cells;
    expect(cells[0]).toMatchObject({ type: "number", value: "123." });
    expect(cells[1]).toMatchObject({ type: "number", value: "4.5e2" });
    expect(cells[2]).toMatchObject({ type: "number", value: "6." });
    // 非数字形态仍留在文本类
    const wb2 = csvToWorkbook(strToU8("1.2.3,.,-\n"));
    const cells2 = wb2.sheets[0].rows[0].cells;
    for (const c of cells2) expect(c.type).toBe("string");
  });
  it("GB18030 回退", () => {
    // "中文,测试" 的 GBK 编码（非法 UTF-8 序列）
    const gbk = new Uint8Array([
      0xd6, 0xd0, 0xce, 0xc4, 0x2c, 0xb2, 0xe2, 0xca, 0xd4, 0x0a,
    ]);
    const wb = csvToWorkbook(gbk);
    expect(wb.sheets[0].rows[0].cells.map((c) => c.value)).toEqual([
      "中文",
      "测试",
    ]);
  });
});
