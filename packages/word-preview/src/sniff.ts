/**
 * 输入形态嗅探：在调用 docx-preview **之前**归一错误。
 *
 * 必要性（实测）：zip 但非 docx（或缺 OPC 关系文件）会让上游崩出模糊的
 * `Cannot read properties of null (reading 'find')`；非 zip 是 jszip 的
 * `Can't find end of central directory`。魔数只能区分 zip/OLE2，区分不了
 * zip 里装的是 docx、xlsx 还是 odt——必须再做 zip 目录的必需部件校验。
 *
 * 主部件路径不能硬编码 word/document.xml：OPC 允许任意主部件路径，由
 * 根关系 `_rels/.rels` 的 officeDocument 关系决定；上游按该关系解析。
 * 注意两侧并非严格同口径：本层的路径归一（反斜杠→斜杠、去 `./` 前缀）
 * 比引擎侧（只去前导 `/`）更宽松，`Target="./word/document.xml"` 这类
 * OPC 合法包会过嗅探、在引擎侧落空、以 CORRUPT 收场——方向是"宁可放过
 * 归 CORRUPT"而非误拒（误拒会把引擎可渲染的文档挡在 UNSUPPORTED）。
 */
import type { WordPreviewErrorCode } from "./types";
import JSZip from "jszip";

export type WordSniffResult =
  | { ok: true }
  | {
      ok: false;
      error: { code: WordPreviewErrorCode; message: string; cause?: unknown };
    };

const CONTENT_TYPES = "[Content_Types].xml";
const ROOT_RELS = "_rels/.rels";
const DEFAULT_MAIN_PART = "word/document.xml";
const OFFICE_DOCUMENT_REL_SUFFIX = "/officeDocument";

function fail(
  code: WordPreviewErrorCode,
  message: string,
  cause?: unknown,
): WordSniffResult {
  return cause === undefined
    ? { ok: false, error: { code, message } }
    : { ok: false, error: { code, message, cause } };
}

/** 十进制 / 十六进制字符引用 → 字符串（非法码点回退替换符）。 */
function decodeCharRef(digits: string, radix: number): string {
  const code = Number.parseInt(digits, radix);
  return code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)
    ? "�"
    : String.fromCodePoint(code);
}

/** XML 属性值实体解码（关系 Target 可能含 &amp; 等）。与读取侧（DOM
 * DOMParser）同口径：十进制与十六进制字符引用都要解，否则合法的
 * `word/&#x64;ocument.xml` 会在本层被误判为缺部件。 */
function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => decodeCharRef(hex, 16))
    .replace(/&#(\d+);/g, (_, dec: string) => decodeCharRef(dec, 10))
    .replace(/&amp;/g, "&");
}

/** 读取标签内的属性（双引号/单引号均可）。 */
function attrValue(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\b${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
  const raw = m?.[1] ?? m?.[2];
  return raw === undefined ? undefined : decodeXmlEntities(raw);
}

/** 关系 Target → 包内部件路径；外部 URI 返回 null。 */
function normalizePartPath(target: string): string | null {
  const value = target.trim();
  if (/^[a-z][a-z0-9+.-]*:/i.test(value)) return null; // 外部 URI
  const path = value
    .replace(/\\/g, "/")
    .replace(/^\/+/, "")
    .replace(/^\.\/+/, "");
  return path === "" ? null : path;
}

function findPart(zip: JSZip, path: string): JSZip.JSZipObject | null {
  return zip.files[path] ?? zip.files[path.replace(/\//g, "\\")] ?? null;
}

/** WordprocessingML 主部件允许的 ContentType（docx / docm）。 */
const WORD_MAIN_CONTENT_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  "application/vnd.ms-word.document.macroenabled.main+xml",
]);

/**
 * 主部件在 [Content_Types].xml 中声明的 ContentType（无 Override 时 null）。
 *
 * 只校验 OPC 部件齐全无法区分 docx / xlsx / pptx——三者都有
 * [Content_Types].xml、_rels/.rels 与 officeDocument 主部件，随后会被
 * 引擎按 Word 解析并失败，用户看到误导性的 CORRUPT。主部件的 Override
 * 是 OPC 里唯一稳定的类型标识。
 */
async function mainPartContentType(
  zip: JSZip,
  mainPart: string,
): Promise<string | null> {
  const ct = findPart(zip, CONTENT_TYPES);
  if (!ct) return null;
  let xml: string;
  try {
    xml = await ct.async("string");
  } catch {
    return null; // 损坏的 Content_Types：交给后续渲染阶段的 CORRUPT 归一
  }
  const target = mainPart.replace(/\\/g, "/").replace(/^\/+/, "");
  for (const tag of xml.match(/<Override\b[^>]*?(?:\/>|>)/g) ?? []) {
    const partName = attrValue(tag, "PartName");
    const contentType = attrValue(tag, "ContentType");
    if (partName === undefined || contentType === undefined) continue;
    if (partName.replace(/\\/g, "/").replace(/^\/+/, "") === target) {
      return contentType.trim();
    }
  }
  return null;
}

/**
 * 主部件路径：根关系的 officeDocument Target；缺失/不可读时回退约定路径。
 * 读取失败不在此报错——必需部件检查会把缺失的 _rels/.rels 或主部件
 * 如实列进 UNSUPPORTED message。
 */
async function resolveMainPart(zip: JSZip): Promise<string> {
  const rels = findPart(zip, ROOT_RELS);
  if (!rels) return DEFAULT_MAIN_PART;
  try {
    const xml = await rels.async("string");
    for (const tag of xml.match(/<Relationship\b[^>]*?(?:\/>|>)/g) ?? []) {
      const type = attrValue(tag, "Type");
      if (!type || !type.endsWith(OFFICE_DOCUMENT_REL_SUFFIX)) continue;
      const target = attrValue(tag, "Target");
      const path = target ? normalizePartPath(target) : null;
      if (path) return path;
    }
  } catch {
    // 关系文件损坏：交给渲染阶段的 CORRUPT 归一
  }
  return DEFAULT_MAIN_PART;
}

/** ZIP 签名：普通条目 / 空包 / 分卷。只认两字节 PK 会把 "PKG" 文本当 zip。 */
function isZipSignature(bytes: Uint8Array): boolean {
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    return false;
  }
  const [a, b] = [bytes[2], bytes[3]];
  return (
    (a === 0x03 && b === 0x04) ||
    (a === 0x05 && b === 0x06) ||
    (a === 0x07 && b === 0x08)
  );
}

export async function sniffWordDocument(
  bytes: Uint8Array,
): Promise<WordSniffResult> {
  if (bytes.length === 0) {
    return fail("UNSUPPORTED", "File is empty.");
  }
  // OLE2 复合文档魔数：.doc 老格式与 OOXML 标准加密（Agile AES）的 docx
  // 整体都是 OLE2、魔数相同，按魔数不可区分；首版两者均不支持，合并表述。
  if (
    bytes.length >= 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1
  ) {
    return fail(
      "LEGACY_FORMAT",
      "Not supported: legacy .doc format and encrypted .docx files are both OLE2 containers and neither is supported.",
    );
  }
  // docx 必为 zip 容器（完整 ZIP 签名）
  if (!isZipSignature(bytes)) {
    return fail(
      "UNSUPPORTED",
      "Not a Word document: expected a .docx (ZIP) file.",
    );
  }
  // zip 加密位：general purpose bit flag 位于首个 local file header 的
  // 偏移 6（小端）的 bit 0。jszip 的 loadAsync 只解析中央目录（不解压
  // 内容），不会报加密——必须显式嗅探（覆盖罕见的直接 zip 加密形态；
  // 主流"加密 docx"是 OLE2 形态，已在上面 LEGACY_FORMAT 分支拦截）。
  if (
    bytes[2] === 0x03 &&
    bytes[3] === 0x04 &&
    bytes.length > 6 &&
    // bytes[6] 的存在性由前面的 length > 6 保证；?? 0 仅为通过
    // noUncheckedIndexedAccess（严格模式下字节索引访问须显式收窄）
    ((bytes[6] ?? 0) & 0x01) === 0x01
  ) {
    return fail(
      "PASSWORD_PROTECTED",
      "Password-protected documents are not supported.",
    );
  }
  return sniffZipParts(bytes);
}

/** zip 目录校验：结构可解析 + 必需部件齐全。loadAsync 是惰性的（只读
 * 中央目录元数据、不解压内容），开销与文件数成正比、与文件大小无关。 */
async function sniffZipParts(bytes: Uint8Array): Promise<WordSniffResult> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    // 中央目录/条目带加密位时 JSZip 抛 "Encrypted zip are not supported"；
    // 这是密码保护，不是结构损坏（旧实现统一归 CORRUPT）。
    if (/encrypted zip/i.test(message)) {
      return fail(
        "PASSWORD_PROTECTED",
        "Password-protected documents are not supported.",
        e,
      );
    }
    // 截断 zip / 中央目录损坏等（实测 jszip 报 Can't find end of central
    // directory 一类错误）——归 CORRUPT 而非 UNSUPPORTED：zip 意图存在，
    // 是结构坏了
    return fail("CORRUPT", "Corrupted file: broken ZIP structure.", e);
  }
  const mainPart = await resolveMainPart(zip);
  const missing = [CONTENT_TYPES, ROOT_RELS, mainPart].filter(
    (p) => findPart(zip, p) === null,
  );
  if (missing.length > 0) {
    return fail(
      "UNSUPPORTED",
      `Not a Word document: missing required part(s) ${missing.map((p) => `"${p}"`).join(", ")}.`,
    );
  }
  const contentType = await mainPartContentType(zip, mainPart);
  if (
    contentType !== null &&
    !WORD_MAIN_CONTENT_TYPES.has(contentType.toLowerCase())
  ) {
    return fail(
      "UNSUPPORTED",
      `Not a Word document: the package main part "${mainPart}" is declared as "${contentType}", not WordprocessingML.`,
    );
  }
  return { ok: true };
}
