/**
 * 输入形态嗅探：只做有确定依据的前置分流，其余交给引擎。
 *
 * 分流规则（docs/ spike 记录 S2/S3 实测上游行为支撑）：
 * 1. 空输入 → UNSUPPORTED（引擎对空输入报 not-ooxml，提前归一更直接）；
 * 2. OLE2/CFB 魔数（D0 CF 11 E0 A1 B1 1A E1）→ **放行给引擎**：它可能是
 *    老式 .ppt、Agile 加密 .pptx 或其他 CFB，三者只能由引擎按 CFB 目录
 *    区分（含 EncryptionInfo → 加密；含 PowerPoint Document 等 → legacy；
 *    其他 → cfb-unknown）。前置阶段断言 legacy 会把加密文件误报成 .ppt；
 * 3. ZIP 魔数（PK\x03\x04 等）→ 放行给引擎（是否缺部件、central directory
 *    是否可读都由引擎判，不为此引入 zip 解析依赖）；
 * 4. 其他 → UNSUPPORTED（既非 OOXML 容器也非 OLE2，没有更多可说的）。
 *
 * 不承诺 password 解 ZIP 条目加密（GP bit 检测首版不做，避免把可打开的
 * 文件误报为不支持）。
 */
import type { PptPreviewErrorCode } from "./types";

export type PptSniffResult =
  | { ok: true }
  | {
      ok: false;
      error: { code: PptPreviewErrorCode; message: string; cause?: unknown };
    };

/** OLE2/CFB 签名（[MS-CFB] 头 8 字节）。 */
const OLE2_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;
/** ZIP local file header 等常见签名前缀（PK）。 */
const ZIP_PREFIX = [0x50, 0x4b] as const;

function startsWith(bytes: Uint8Array, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) return false;
  for (let i = 0; i < prefix.length; i++) {
    if (bytes[i] !== prefix[i]) return false;
  }
  return true;
}

export function sniffPpt(bytes: Uint8Array): PptSniffResult {
  if (bytes.length === 0) {
    return {
      ok: false,
      error: { code: "UNSUPPORTED", message: "File is empty." },
    };
  }
  if (startsWith(bytes, OLE2_MAGIC)) return { ok: true };
  if (startsWith(bytes, ZIP_PREFIX)) return { ok: true };
  return {
    ok: false,
    error: {
      code: "UNSUPPORTED",
      message:
        "Not a PowerPoint file: the input is neither a ZIP-based .pptx package nor an OLE2 container (legacy .ppt or encrypted .pptx).",
    },
  };
}
