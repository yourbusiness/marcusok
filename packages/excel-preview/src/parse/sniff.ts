/**
 * 输入形态嗅探：xlsx(zip) / OLE2(.xls) / 纯文本(csv) / 未知。
 * worker 与主线程回退路径共用（worker 构建会把本模块打入单文件产物，
 * 此前两侧各持一份拷贝）。
 */
export type SniffedFormat = "zip" | "ole2" | "text" | "unknown";

/**
 * 采样段能否被 UTF-8 或 GB18030 以 fatal 语义完整解码（与 csv.ts 的
 * decodeBytes 解码链同集合）。可打印率计数对 ≥0x80 字节宽容（多字节编码），
 * 但连续 0xFF 一类的字节流不是任何合法编码——纯二进制（图标/字体/加密块）
 * 的可打印率可以虚高到 95% 阈值之上，此前会被当 CSV 渲染成乱码而非
 * UNSUPPORTED 报错。必须真正可解码才算文本。
 *
 * 截断兜底：采样截在多字节序列中间（2KB 边界切在 GB18030 双字节的中间）时
 * fatal 解码必然失败——去掉尾部 4 字节（覆盖 GB18030 最长四字节序列）重试
 * 一次，仍失败才否定。
 */
function looksDecodable(sample: Uint8Array): boolean {
  const attempts = [sample, sample.subarray(0, Math.max(0, sample.length - 4))];
  for (const bytes of attempts) {
    if (bytes.length === 0) continue;
    if (canDecode(bytes, "utf-8") || canDecode(bytes, "gb18030")) return true;
  }
  return false;
}

function canDecode(bytes: Uint8Array, encoding: string): boolean {
  try {
    new TextDecoder(encoding, { fatal: true }).decode(bytes);
    return true;
  } catch {
    return false;
  }
}

export function sniffFormat(bytes: Uint8Array): SniffedFormat {
  // 只认完整的 ZIP 签名（普通条目 PK\x03\x04、空包 PK\x05\x06、分卷
  // PK\x07\x08）。仅匹配前两字节 "PK" 会把以 "PK," / "PKG" 开头的合法
  // CSV（主键列很常见的表头）误判成 zip，最终报出 "Legacy .xls" 的错误。
  if (
    bytes.length >= 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    ((bytes[2] === 0x03 && bytes[3] === 0x04) ||
      (bytes[2] === 0x05 && bytes[3] === 0x06) ||
      (bytes[2] === 0x07 && bytes[3] === 0x08))
  ) {
    return "zip";
  }
  // OLE2 复合文档签名是完整 8 字节；截断的 4 字节前缀不应命中。
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
    return "ole2";
  }
  // 纯文本嗅探：前 2KB 中可打印字符占比（文本文件常态 ≥ 95%）
  const n = Math.min(bytes.length, 2048);
  if (n === 0) return "unknown";
  let printable = 0;
  for (let i = 0; i < n; i++) {
    const b = bytes[i];
    if (b === 9 || b === 10 || b === 13 || (b >= 0x20 && b <= 0x7e))
      printable++;
    else if (b >= 0x80) printable++; // 多字节编码按可打印宽容处理
  }
  if (printable / n < 0.95) return "unknown";
  // 可打印 ≠ CSV：SpreadsheetML 2003（<?xml 开头的伪 .xls）与 HTML 表格
  // 另存的 .xls 也是高可打印文本，此前会被当 CSV 渲染成标签碎行（与文档
  // "非 ZIP → CORRUPT"的声明不符）。按 unknown 归入 UNSUPPORTED 友好报错。
  let i = 0;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) i = 3; // UTF-8 BOM
  let head = "";
  for (; i < Math.min(bytes.length, 256); i++) {
    head += String.fromCharCode(bytes[i]);
  }
  head = head.trimStart().toLowerCase();
  if (
    head.startsWith("<?xml") ||
    head.startsWith("<!doctype") ||
    head.startsWith("<html")
  ) {
    return "unknown";
  }
  // 高可打印还须真可解码：0x80+ 字节占比虚高的二进制（如 0xFF 填充流，
  // 连续 0xFF 不属于 UTF-8 也不是合法 GB18030 双/四字节的任何位置）在
  // 两条 fatal 解码下都失败，归 unknown → UNSUPPORTED。
  if (!looksDecodable(bytes.subarray(0, n))) return "unknown";
  return "text";
}
