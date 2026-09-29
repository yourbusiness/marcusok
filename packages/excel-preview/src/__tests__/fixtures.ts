/**
 * 测试样本构造：手工 OOXML（fflate 打 zip）+ modern-xlsx 写 API 两条路。
 *
 * 手工 OOXML 的意义：modern-xlsx 的写 API 造不出主题色（theme+tint）与
 * 特定形态（date1904、指定 view 等），而它们正是预览层要还原的关键路径
 * （实测依据见 styles-overlay.ts 注释）。
 */
import { strToU8, zipSync } from "fflate";
import { initWasmSync, Workbook } from "@marcusok/xlsx-core";

// Node 测试环境 wasm 引导（与 excel-exporter setup.ts 同一做法；经包内
// devDep 的 modern-xlsx 解析，与 core 共享同一份二进制）。注意用字符串路径
// 而非 new URL(...)：happy-dom 环境覆写了全局 URL，node:fs 不认其实例
// （实测报 "The URL must be of scheme file"）。
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
initWasmSync(
  readFileSync(
    join(
      dirname(fileURLToPath(import.meta.url)),
      "../../node_modules/modern-xlsx/dist/modern-xlsx.wasm",
    ),
  ),
);

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const x = (s: string) => strToU8(HEAD + s);

export const OFFICE_THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`;

export interface SheetSpec {
  name: string;
  /** sheetData 的行（1-based r 属性自动生成）。 */
  xml: string;
  /** sheet 根上的额外元素（cols/mergeCells/sheetViews 挂 sheetData 前后按需拼）。 */
  extraBefore?: string;
  extraAfter?: string;
  state?: "visible" | "hidden" | "veryHidden";
  /** sheet 级 rels（hyperlink 等需要）。 */
  rels?: Record<string, string>;
}

export interface XlsxSpec {
  sheets: SheetSpec[];
  styles: string;
  sharedStrings?: string;
  theme?: string;
  date1904?: boolean;
  activeTab?: number;
}

/** 默认 styles：2 字体（font1 带 theme 色）+ 主题/ indexed 填充 + 主题色边框 + numFmt。
 * border#0 用自闭合 <border/>——真实 Excel 产物的常态形态（Excel 对默认
 * 无边框项写自闭合标签；modern-xlsx 自家写产物是展开形态，若不用自闭合
 * 造样，引擎跳过自闭合导致的 id 错位在测试里永远踩不到）。 */
export const RICH_STYLES = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="¥#,##0.00"/></numFmts><fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color theme="1" tint="-0.25"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font><font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.6"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor indexed="44"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border/><border><left/><right/><top/><bottom style="thin"><color theme="4"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="8"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1" indent="1"/></xf><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;

/** 构造完整 xlsx zip 字节。 */
export function buildXlsx(spec: XlsxSpec): Uint8Array {
  const overrides = spec.sheets
    .map(
      (_, i) =>
        `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join("");
  const ct = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${overrides}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${spec.sharedStrings ? `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` : ""}</Types>`;

  const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

  const datePr = spec.date1904 ? `<workbookPr date1904="1"/>` : "";
  const views =
    spec.activeTab !== undefined
      ? `<bookViews><workbookView activeTab="${spec.activeTab}"/></bookViews>`
      : "";
  const wb = `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${datePr}${views}<sheets>${spec.sheets
    .map(
      (s, i) =>
        `<sheet name="${s.name}" sheetId="${i + 1}"${s.state && s.state !== "visible" ? ` state="${s.state}"` : ""} r:id="rId${i + 1}"/>`,
    )
    .join("")}</sheets></workbook>`;

  const wbRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${spec.sheets
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
    )
    .join(
      "",
    )}<Relationship Id="rIdW" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>${spec.sharedStrings ? `<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` : ""}</Relationships>`;

  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": x(ct),
    "_rels/.rels": x(rels),
    "xl/workbook.xml": x(wb),
    "xl/_rels/workbook.xml.rels": x(wbRels),
    "xl/styles.xml": x(spec.styles),
    "xl/theme/theme1.xml": x(spec.theme ?? OFFICE_THEME),
  };
  if (spec.sharedStrings) {
    files["xl/sharedStrings.xml"] = x(spec.sharedStrings);
  }
  spec.sheets.forEach((s, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = x(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${s.extraBefore ?? ""}<sheetData>${s.xml}</sheetData>${s.extraAfter ?? ""}</worksheet>`,
    );
    if (s.rels) {
      const entries = Object.entries(s.rels)
        .map(
          ([id, target]) =>
            `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${target}" TargetMode="External"/>`,
        )
        .join("");
      files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = x(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`,
      );
    }
  });
  return zipSync(files);
}

/** 基础样本：主题色字体/填充、indexed 填充、合并、冻结、隐藏行列、公式缓存。 */
export function sampleWorkbookBytes(): Uint8Array {
  return buildXlsx({
    sheets: [
      {
        name: "S1",
        extraBefore:
          '<cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="8.43" hidden="1"/></cols><sheetViews><sheetView tabSelected="1" workbookViewId="0" showGridLines="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>',
        xml: `<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>themeFont</t></is></c><c r="B1" s="2" t="inlineStr"><is><t>themeFill</t></is></c><c r="C1" s="3" t="inlineStr"><is><t>idxFill</t></is></c><c r="D1" s="6" t="inlineStr"><is><t>rgbFill</t></is></c></row><row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>hiddenRow</t></is></c></row><row r="3"><c r="A3"><f>1+2</f><v>3</v></c><c r="B3" t="b"><v>1</v></c><c r="C3" t="e"><v>#DIV/0!</v></c><c r="D3"><v>1234.5</v></c><c r="E3" s="5"><v>1234.5</v></c><c r="F3" s="4"><v>45678</v></c></row><row r="4" ht="30" customHeight="1"><c r="A4" t="inlineStr"><is><t>tall</t></is></c></row>`,
        extraAfter:
          '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
      },
      {
        name: "S2",
        xml: `<row r="1"><c r="A1" t="inlineStr"><is><t>s2</t></is></c></row>`,
      },
      {
        name: "Hidden",
        xml: `<row r="1"><c r="A1" t="inlineStr"><is><t>h</t></is></c></row>`,
        state: "hidden",
      },
    ],
    styles: RICH_STYLES,
    activeTab: 0,
  });
}

/** date1904 样本：serial 0 = 1904-01-01。 */
export function date1904Bytes(): Uint8Array {
  return buildXlsx({
    sheets: [
      {
        name: "D4",
        xml: `<row r="1"><c r="A1" s="4"><v>0</v></c><c r="B1" s="4"><v>1.5</v></c></row>`,
      },
    ],
    styles: RICH_STYLES,
    date1904: true,
  });
}

/** 自闭合样式元素专项样本（真实 Excel 产物形态）：font#0/fill#0/border#0
 * 全部自闭合。回归锚点：引擎读取侧跳过自闭合元素，其 fonts/fills/borders
 * 数组各少一项、与 cellXfs 的 id 引用整体错位（borderId=0 拿到文件 border#1
 * 的样式、末位引用越界丢失）——模型必须经覆盖层重建为含空项的完整列表。 */
export function selfClosedStylesBytes(): Uint8Array {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill/><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="3"><border/><border><left/><right/><top/><bottom style="thin"><color theme="1"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="medium"><color rgb="FFFF0000"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/></cellXfs></styleSheet>`;
  return buildXlsx({
    sheets: [
      {
        name: "SC",
        // A1(xf1)=thin 黑底边；B1(xf2)=medium 红底边；C1(xf3)=粗体；D1(xf4)=黄填充
        xml: `<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>a</t></is></c><c r="B1" s="2" t="inlineStr"><is><t>b</t></is></c><c r="C1" s="3" t="inlineStr"><is><t>c</t></is></c><c r="D1" s="4" t="inlineStr"><is><t>d</t></is></c></row>`,
      },
    ],
    styles,
  });
}

/** sharedStrings 富文本样本。 */
export function richTextBytes(): Uint8Array {
  return buildXlsx({
    sheets: [
      {
        name: "RT",
        xml: `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>`,
      },
    ],
    styles: RICH_STYLES,
    sharedStrings:
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Hel</t></r><r><rPr><color rgb="FF0000FF"/></rPr><t>lo</t></r></si><si><t>plain</t></si></sst>',
  });
}

/** 伪造 OLE2 头（.xls 判别路径）。 */
export function fakeOle2Bytes(): Uint8Array {
  const b = new Uint8Array(512);
  b.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  return b;
}

/** modern-xlsx 写 API 造加密文件。 */
export async function encryptedBytes(password: string): Promise<Uint8Array> {
  const wb = new Workbook();
  wb.addSheet("E").cell("A1").value = "secret";
  return wb.toBuffer({ password });
}

/** 大样本（性能基线）：rows 行 × cols 列数字。 */
export function bigWorkbookBytes(rows: number, cols: number): Uint8Array {
  const chunk: string[] = [];
  for (let r = 1; r <= rows; r++) {
    const cells: string[] = [];
    for (let c = 1; c <= cols; c++) {
      const ref = colName(c) + r;
      cells.push(`<c r="${ref}"><v>${r * cols + c}</v></c>`);
    }
    chunk.push(`<row r="${r}">${cells.join("")}</row>`);
  }
  return buildXlsx({
    sheets: [{ name: "Big", xml: chunk.join("") }],
    styles: `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`,
  });
}

export function colName(c: number): string {
  let s = "";
  let n = c;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 读取入口（wasm 已在本文件头部初始化）。 */
export { readBuffer } from "@marcusok/xlsx-core";
