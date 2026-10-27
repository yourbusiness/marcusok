/**
 * sniff 纯字节测试（node 环境，不触 pdf.js）。
 *
 * 口径：%PDF- 魔数在头部 1024 字节窗口内搜索（Adobe 实现口径，允许
 * 文件头前有垃圾字节）；窗口外属有意的输入收紧（limits 已注明）。
 */
import { describe, it, expect } from "vitest";
import { sniffPdf } from "../sniff";
import { ZIP_MAGIC_BYTES, makePdfBytes, withJunkPrefix } from "./fixtures";

describe("sniffPdf", () => {
  it("accepts a well-formed PDF (magic at offset 0)", async () => {
    const bytes = await makePdfBytes({ pages: 1 });
    expect(sniffPdf(bytes)).toEqual({ ok: true });
  });

  it("accepts a PDF with junk bytes before the header (within window)", async () => {
    const bytes = await makePdfBytes({ pages: 1 });
    expect(sniffPdf(withJunkPrefix(bytes, 512))).toEqual({ ok: true });
  });

  it("rejects a PDF whose header sits beyond the 1024-byte window", async () => {
    const bytes = await makePdfBytes({ pages: 1 });
    const result = sniffPdf(withJunkPrefix(bytes, 1200));
    expect(result).toEqual({
      ok: false,
      error: {
        code: "UNSUPPORTED",
        message:
          "Not a PDF: the %PDF- header was not found within the first 1024 bytes.",
      },
    });
  });

  it("rejects empty input", () => {
    expect(sniffPdf(new Uint8Array(0))).toEqual({
      ok: false,
      error: { code: "UNSUPPORTED", message: "File is empty." },
    });
  });

  it("rejects zip magic (xlsx/docx-shaped input)", () => {
    const result = sniffPdf(ZIP_MAGIC_BYTES);
    expect(result.ok).toBe(false);
    expect(result.ok === false && result.error.code).toBe("UNSUPPORTED");
  });

  it("rejects truncated input shorter than the magic", () => {
    // "%PD" 截断：窗口内凑不出完整 %PDF-
    const short = Uint8Array.from([0x25, 0x50, 0x44]);
    expect(sniffPdf(short).ok).toBe(false);
  });

  it("does not match %PDF without the trailing dash", () => {
    // "%PDFX" 等近形态不得误判
    const near = Uint8Array.from([0x25, 0x50, 0x44, 0x46, 0x58, 0x0a]);
    expect(sniffPdf(near).ok).toBe(false);
  });
});
