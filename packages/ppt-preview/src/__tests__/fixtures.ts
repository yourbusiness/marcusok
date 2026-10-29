/**
 * 测试样本构造：fflate 打 zip 手工构造最小 pptx（word/excel-preview 的
 * fixture 模式；最小配方经上游实测收敛——无需 slideMaster/slideLayout/
 * theme，四类部件即可，见 docs/ppt-preview-spike-record.md S3）。
 *
 * 加密样本走 agile-crypto.ts（cfb + Node crypto 生成真实 MS-OFFCRYPTO
 * Agile 容器，闭环验证过上游可解密），不用 mock 冒充加密验收。
 */
import { strToU8, zipSync } from "fflate";

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const u = (s: string) => strToU8(HEAD + s);

export interface PptxSpec {
  /** 每页标题文本（页数 = texts.length）。 */
  texts: string[];
  /** 幻灯片尺寸（EMU，默认 12192000x6858000 = 16:9）。 */
  size?: { cx: number; cy: number };
  /** 额外的部件覆盖（如缺件、跨格式内容），键为 zip 路径。 */
  overrides?: Record<string, Uint8Array>;
  /** 从默认集合中删除的部件路径。 */
  omit?: string[];
}

const SLIDE = (text: string) =>
  `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="4400"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

/** 构造完整 pptx zip 字节（texts 为空数组 = 空 deck）。 */
export function buildPptx(spec: PptxSpec): Uint8Array {
  const n = spec.texts.length;
  const { cx = 12192000, cy = 6858000 } = spec.size ?? {};
  const slideOverrides = Object.fromEntries(
    Array.from({ length: n }, (_, i) => [
      `ppt/slides/slide${i + 1}.xml`,
      u(SLIDE(spec.texts[i])),
    ]),
  );
  const ctEntries = [
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`,
    ...Array.from(
      { length: n },
      (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
    ),
  ].join("");
  const pres = `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst>${Array.from(
    { length: n },
    (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`,
  ).join(
    "",
  )}</p:sldIdLst><p:sldSz cx="${cx}" cy="${cy}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  const presRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${Array.from(
    { length: n },
    (_, i) =>
      `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
  ).join("")}</Relationships>`;

  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": u(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${ctEntries}</Types>`,
    ),
    "_rels/.rels": u(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    ),
    "ppt/presentation.xml": u(pres),
    "ppt/_rels/presentation.xml.rels": u(presRels),
    ...slideOverrides,
    ...(spec.overrides ?? {}),
  };
  for (const key of spec.omit ?? []) delete files[key];
  return zipSync(files);
}

/** 单页便捷构造。 */
export function simplePptx(text = "Hello ppt"): Uint8Array {
  return buildPptx({ texts: [text] });
}

/** zip 但非 pptx：单文件普通 zip（无任何 OPC 部件）。 */
export function plainZipBytes(): Uint8Array {
  return zipSync({ "hello.txt": strToU8("not a pptx") });
}

/** 跨格式：最小 docx（上游报 not-ooxml 且 message 指明缺 ppt/presentation.xml）。 */
export function docxBytes(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": u(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": u(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": u(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>`,
    ),
  });
}

/** 伪造 OLE2 头（512B 魔数 + 零填充，无合法 CFB 目录——引擎归 not-ooxml）。 */
export function fakeOle2Bytes(): Uint8Array {
  const b = new Uint8Array(512);
  b.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  return b;
}
