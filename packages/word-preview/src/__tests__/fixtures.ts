/**
 * 测试样本构造：fflate 打 zip 手工构造最小 docx（不引入二进制样本文件，
 * 与 excel-preview 的 fixture 模式一致）。
 *
 * 最小可行集合是**四文件**：`[Content_Types].xml`、`_rels/.rels`（根关系，
 * 指向 word/document.xml）、`word/document.xml`、`word/_rels/document.xml.rels`
 * （可为空 Relationships）。两文件/三文件集合会让 docx-preview 崩出与
 * "zip 非 docx"完全相同的模糊错误（`Cannot read properties of null
 * (reading 'find')`，Node + happy-dom spike 实测），别按错误配方写。
 * Content_Types 需含 `.rels` 扩展的 Default 声明。
 */
import { strToU8, unzipSync, zipSync } from "fflate";

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const u = (s: string) => strToU8(HEAD + s);

const CONTENT_TYPES = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;

const ROOT_RELS = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const EMPTY_DOC_RELS = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

export interface DocxSpec {
  /** word/document.xml 的 <w:body> 内文（段落/表格/分节符的原始 XML）。 */
  body: string;
  /** word/_rels/document.xml.rels 内容（缺省空 Relationships；超链接等
   * 关系类 fixture 需自定义）。 */
  docRels?: string;
}

/** 构造完整 docx zip 字节。body 必须显式给出（多段落/表格等组合 XML）。 */
export function buildDocx(spec: DocxSpec): Uint8Array {
  const doc = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${spec.body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": u(CONTENT_TYPES),
    "_rels/.rels": u(ROOT_RELS),
    "word/document.xml": u(doc),
    "word/_rels/document.xml.rels": u(spec.docRels ?? EMPTY_DOC_RELS),
  });
}

/** 单段文本便捷构造（内部即 buildDocx({ body: parasXml(text) })）。 */
export function simpleDocx(text: string): Uint8Array {
  return buildDocx({ body: parasXml(text) });
}

/** 纯文本段落的快捷构造（不含 XML 转义，测试文本自控）。 */
export function parasXml(text: string): string {
  return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

/** 多段落样本（每段一个 <w:p>）。 */
export function parasListXml(texts: string[]): string {
  return texts.map((t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`).join("");
}

/**
 * 分页样本：两段文本之间一个 `<w:lastRenderedPageBreak/>`（元素名而非
 * 带前缀属性，happy-dom 能解析——`<w:br w:type="page"/>` 的 w:type 是
 * 带前缀属性，happy-dom 下上游 attr() 读不到、不会分页，真实分页符
 * 的复核在浏览器 demo）。本包默认 `ignoreLastRenderedPageBreak: false`
 * （偏离上游默认 true），该标记即触发断页。
 */
export function pageBreakBetween(a: string, b: string): string {
  return `<w:p><w:r><w:t>${a}</w:t></w:r></w:p><w:p><w:r><w:lastRenderedPageBreak/><w:t>${b}</w:t></w:r></w:p>`;
}

/** 伪造 OLE2 头（.doc / 加密 docx 判别路径）。 */
export function fakeOle2Bytes(): Uint8Array {
  const b = new Uint8Array(512);
  b.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  return b;
}

/** zip 但非 docx：单文件普通 zip（有 zip 结构、无任何 OPC 部件）。 */
export function plainZipBytes(): Uint8Array {
  return zipSync({ "hello.txt": strToU8("not a docx") });
}

/** 缺特定部件的 docx：从完整四文件集合里删掉指定键。 */
export function docxMissingPart(part: string): Uint8Array {
  const full = simpleDocx("x");
  // 重新组包（zip 不可编辑）：解开再剔除
  const files = unzipSync(full);
  const out: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) {
    if (k !== part) out[k] = v;
  }
  return zipSync(out);
}

/**
 * 手工构造带 general purpose bit 0（加密位）的 zip 头 fixture：只需让嗅探
 * 在字节层看到该位——复用完整 docx 字节，把首个 local file header 的
 * GP flag（偏移 6，小端）置 bit 0。文件名长度/数据在 header 之后，改
 * 一个 flag 字节不破坏 zip 结构布局（jszip 读中央目录，目录里的 GP 位
 * 与 local header 不一致不影响 loadAsync 成败——嗅探发生在 loadAsync 之前）。
 */
export function encryptedZipBytes(): Uint8Array {
  const bytes = new Uint8Array(simpleDocx("secret"));
  bytes[6] |= 0x01;
  return bytes;
}

/** 大样本（性能基线）：纯文本——1000 段落 + 一个 100 行 × 5 列表格。
 * 不放图片/内嵌字体：renderAsync resolve 前 await 全部资源任务，放图会把
 * base64 编码耗时计入，"DOM 构建"口径即失真。 */
export function bigDocxBytes(
  paragraphs: number,
  tableRows: number,
  tableCols: number,
): Uint8Array {
  const paras: string[] = [];
  for (let i = 0; i < paragraphs; i++) {
    paras.push(
      `<w:p><w:r><w:t>Paragraph ${i} — 中英文混排 benchmark text for word preview rendering pipeline.</w:t></w:r></w:p>`,
    );
  }
  const rows: string[] = [];
  for (let r = 0; r < tableRows; r++) {
    const cells: string[] = [];
    for (let c = 0; c < tableCols; c++) {
      cells.push(
        `<w:tc><w:tcPr/><w:p><w:r><w:t>cell ${r}-${c}</w:t></w:r></w:p></w:tc>`,
      );
    }
    rows.push(`<w:tr>${cells.join("")}</w:tr>`);
  }
  const tbl = `<w:tbl><w:tblPr/><w:tblGrid>${"<w:gridCol/>".repeat(tableCols)}</w:tblGrid>${rows.join("")}</w:tbl>`;
  return buildDocx({ body: paras.join("") + tbl });
}
