/**
 * 输入形态嗅探：在调用 docx-preview **之前**归一错误。
 *
 * 必要性（实测）：zip 但非 docx（或缺 OPC 关系文件）会让上游崩出模糊的
 * `Cannot read properties of null (reading 'find')`；非 zip 是 jszip 的
 * `Can't find end of central directory`。魔数只能区分 zip/OLE2，区分不了
 * zip 里装的是 docx、xlsx 还是 odt——必须再做 zip 目录的必需部件校验。
 */
import type { WordPreviewErrorCode } from "./types";
import JSZip from "jszip";

export type WordSniffResult =
  | { ok: true }
  | {
      ok: false;
      error: { code: WordPreviewErrorCode; message: string; cause?: unknown };
    };

/**
 * Word 必需部件（主流约定路径，覆盖 Word/WPS/LibreOffice 产物）。与上游
 * `OpenXmlPackage` 内部取文件的方式一致：正斜杠键查不到时兼容反斜杠键
 * （部分打包器产物的路径形态）。
 */
const REQUIRED_PARTS = [
  "[Content_Types].xml",
  "_rels/.rels",
  "word/document.xml",
] as const;

function fail(
  code: WordPreviewErrorCode,
  message: string,
  cause?: unknown,
): WordSniffResult {
  return cause === undefined
    ? { ok: false, error: { code, message } }
    : { ok: false, error: { code, message, cause } };
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
    bytes[3] === 0xe0
  ) {
    return fail(
      "LEGACY_FORMAT",
      "Not supported: legacy .doc format and encrypted .docx files are both OLE2 containers and neither is supported.",
    );
  }
  // docx 必为 zip 容器（PK 魔数）
  if (!(bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b)) {
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
    (bytes[6] & 0x01) === 0x01
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
    // 截断 zip / 中央目录损坏等（实测 jszip 报 Can't find end of central
    // directory 一类错误）——归 CORRUPT 而非 UNSUPPORTED：zip 意图存在，
    // 是结构坏了
    return fail("CORRUPT", "Corrupted file: broken ZIP structure.", e);
  }
  const missing = REQUIRED_PARTS.filter(
    (p) =>
      zip.files[p] === undefined &&
      zip.files[p.replace(/\//g, "\\")] === undefined,
  );
  if (missing.length > 0) {
    return fail(
      "UNSUPPORTED",
      `Not a Word document: missing required part(s) ${missing.map((p) => `"${p}"`).join(", ")}.`,
    );
  }
  return { ok: true };
}
