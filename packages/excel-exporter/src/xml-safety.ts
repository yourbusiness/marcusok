/**
 * OOXML 的 XML 1.0 字符约束。
 *
 * modern-xlsx 的写 API 会把 sheet 名、font.name、numFormat 等字符串原样拼进
 * XML：控制字符（除 tab/LF/CR）、U+FFFE/U+FFFF 与孤立代理项会产出 Excel
 * 判损的文件，而导出仍返回 success: true。数据/表头文本走同一套清洗，避免
 * Workbook 与 fast stream 两条路径写出不同内容。
 */

/** XML 1.0 Char 产生式不允许的 BMP 码点（代理项单独判定）。 */
export function isInvalidXmlCode(code: number): boolean {
  return (
    (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) ||
    code === 0xfffe ||
    code === 0xffff
  );
}

/** 返回首个非法码点（孤立代理项返回其码元）；全部合法返回 null。 */
export function firstInvalidXmlCode(value: string): number | null {
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        i++; // 合法代理对
        continue;
      }
      return code;
    }
    if (code >= 0xdc00 && code <= 0xdfff) return code;
    if (isInvalidXmlCode(code)) return code;
  }
  return null;
}

/** 错误消息里的码点展示（U+0001 / U+FFFE / U+D800）。 */
export function formatCodePoint(code: number): string {
  return `U+${code.toString(16).toUpperCase().padStart(4, "0")}`;
}

/**
 * 数据/表头文本的清洗：删除 XML 非法控制字符，孤立代理项替换为 U+FFFD。
 * 与 fast-xlsx 的 sanitizeXml / TextEncoder 行为对齐，保证两条导出路径的
 * 单元格内容一致。
 */
export function sanitizeXmlText(value: string): string {
  let out = "";
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += value.charAt(i) + value.charAt(i + 1);
        i++;
      } else {
        out += "�";
      }
      continue;
    }
    if (code >= 0xdc00 && code <= 0xdfff) {
      out += "�";
      continue;
    }
    if (isInvalidXmlCode(code)) continue;
    out += value.charAt(i);
  }
  return out;
}

/**
 * 结构性/配置性字符串（sheet 名、font.name、numFormat、format.pattern）的
 * 前置校验：必须是字符串且不含 XML 非法字符。这些字段无法像数据那样静默
 * 清洗（改表名/字体名/格式码会改变业务语义），因此直接报错。
 */
export function assertXmlSafeText(
  value: unknown,
  at: string,
  field: string,
): asserts value is string {
  if (typeof value !== "string") {
    throw new Error(`[excel-exporter] ${at} ${field} must be a string`);
  }
  const invalid = firstInvalidXmlCode(value);
  if (invalid !== null) {
    throw new Error(
      `[excel-exporter] ${at} ${field} contains ${formatCodePoint(invalid)}, which XML/OOXML does not allow`,
    );
  }
}
