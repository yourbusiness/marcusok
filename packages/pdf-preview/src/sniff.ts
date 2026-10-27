/**
 * 输入形态嗅探：在调用 pdf.js 之前归一"不是 PDF"的错误。
 *
 * Adobe PDF 见范允许文件头之前存在垃圾字节，读者应在前 1024 字节内
 * 搜索 `%PDF-` 魔数（PDF 32000-1 §7.5.2 与 Adobe 自家实现口径）。
 * 超过 1024 字节垃圾前缀的畸形文件会被归入 UNSUPPORTED——比官方
 * viewer 更严格的输入收紧，limits 文档已注明。
 */
import type { PdfPreviewErrorCode } from "./types";

export type PdfSniffResult =
  | { ok: true }
  | {
      ok: false;
      error: { code: PdfPreviewErrorCode; message: string; cause?: unknown };
    };

/** `%PDF-` = 0x25 0x50 0x44 0x46 0x2D */
const MAGIC = [0x25, 0x50, 0x44, 0x46, 0x2d] as const;

/** 文件头魔数搜索窗口（Adobe 实现口径：前 1024 字节）。 */
const MAGIC_WINDOW = 1024;

export function sniffPdf(bytes: Uint8Array): PdfSniffResult {
  if (bytes.length === 0) {
    return {
      ok: false,
      error: { code: "UNSUPPORTED", message: "File is empty." },
    };
  }
  const end = Math.min(MAGIC_WINDOW, bytes.length - MAGIC.length + 1);
  outer: for (let i = 0; i < end; i++) {
    for (let j = 0; j < MAGIC.length; j++) {
      if (bytes[i + j] !== MAGIC[j]) continue outer;
    }
    return { ok: true };
  }
  return {
    ok: false,
    error: {
      code: "UNSUPPORTED",
      message:
        "Not a PDF: the %PDF- header was not found within the first 1024 bytes.",
    },
  };
}
