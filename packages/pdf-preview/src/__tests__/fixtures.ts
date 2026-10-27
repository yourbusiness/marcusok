/**
 * 测试 fixture：pdf-lib 生成的合法 PDF（xref/对象结构正确，手写最小 PDF
 * 的维护成本高）+ 非法输入构造器。
 *
 * 边界记录（见方案风险 10/11）：pdf-lib 不生成加密 PDF（密码路径用
 * mock getDocument 抛 PasswordException 验证归一）；pdf-lib 内置无 CJK
 * 字体，中文 PDF 的 cmap 链路在 play/docs demo 真实浏览器验收，node
 * fixture 只覆盖拉丁文本管线。
 */
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

export interface PdfFixtureOptions {
  pages?: number;
  /** 页面点尺寸，默认 US Letter（612×792）。 */
  size?: [number, number];
  /** 每页文本（页码从 1 起），默认 "Fixture page N"。 */
  text?: (pageNo: number) => string;
}

export async function makePdfBytes(
  opts: PdfFixtureOptions = {},
): Promise<Uint8Array> {
  const { pages = 2, size = [612, 792], text } = opts;
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= pages; i++) {
    const page = doc.addPage(size);
    page.drawText(text ? text(i) : `Fixture page ${i}`, {
      x: 36,
      y: size[1] - 54,
      size: 12,
      font,
      color: rgb(0, 0, 0),
    });
  }
  return doc.save();
}

/** 合法 PDF 前置 N 字节垃圾（Adobe 允许文件头前有垃圾字节）。 */
export function withJunkPrefix(bytes: Uint8Array, n: number): Uint8Array {
  const junk = new Uint8Array(n);
  for (let i = 0; i < junk.length; i++) junk[i] = 0x41 + (i % 26);
  const out = new Uint8Array(junk.length + bytes.length);
  out.set(junk);
  out.set(bytes, junk.length);
  return out;
}

/** 魔数合法但结构损坏：截断到前 n 字节。 */
export function truncated(bytes: Uint8Array, n: number): Uint8Array {
  return bytes.slice(0, n);
}

/** 非 PDF 输入：zip 魔数（xlsx/docx 同形）。 */
export const ZIP_MAGIC_BYTES = Uint8Array.from([
  0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00, 0x08, 0x00,
]);
