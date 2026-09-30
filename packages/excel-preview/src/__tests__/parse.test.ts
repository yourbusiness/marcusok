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
  fakeOle2Bytes,
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
import { strToU8, zipSync } from "fflate";

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
  it("BOM 剥离与数字类型推断", () => {
    const wb = csvToWorkbook(
      new Uint8Array([0xef, 0xbb, 0xbf, 0x31, 0x2c, 0x32, 0x0a]),
    );
    const cells = wb.sheets[0].rows[0].cells;
    expect(cells[0]).toMatchObject({ type: "number", value: "1" });
    expect(cells[1]).toMatchObject({ type: "number", value: "2" });
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
