/**
 * 输入形态嗅探：xlsx(zip) / OLE2(.xls) / 纯文本(csv) / 未知。
 * worker 与主线程回退路径共用（worker 构建会把本模块打入单文件产物，
 * 此前两侧各持一份拷贝）。
 */
export type SniffedFormat = "zip" | "ole2" | "text" | "unknown";

export function sniffFormat(bytes: Uint8Array): SniffedFormat {
  if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b) return "zip";
  if (
    bytes.length >= 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0
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
  return "text";
}
