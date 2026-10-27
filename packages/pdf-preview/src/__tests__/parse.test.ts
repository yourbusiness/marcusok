/**
 * parse 层测试（node 环境 + vitest alias 定向 pdfjs-dist legacy 构建——
 * 现代构建在 Node 下 getDocument 崩，见 vitest.config.ts 注释）。
 *
 * 覆盖两层：
 * 1. 真实引擎管线：pdf-lib fixture → getDocument → numPages/viewport/
 *    textContent（验证 fixture 与上游 API 行为假设）；
 * 2. 错误形态实证：截断/损坏输入的上游异常名（render.test 的错误归一
 *    映射以此为依据——node 直测引擎，映射逻辑本身在 happy-dom 侧经
 *    createPdfPreview 验证）。
 */
import { describe, it, expect } from "vitest";
import { getDocument } from "pdfjs-dist";
import type { PDFDocumentLoadingTask } from "pdfjs-dist";
import { makePdfBytes, STANDARD_FONT_DATA_URL, truncated } from "./fixtures";

describe("pdf.js legacy parse pipeline (real engine)", () => {
  it("parses a 3-page fixture: numPages, per-page viewport, text content", async () => {
    const bytes = await makePdfBytes({ pages: 3 });
    const task: PDFDocumentLoadingTask = getDocument({
      data: bytes,
      standardFontDataUrl: STANDARD_FONT_DATA_URL,
    });
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(3);

    const page = await pdf.getPage(2);
    const vp = page.getViewport({ scale: 1 });
    // fixture 默认 US Letter（612×792pt）
    expect(Math.round(vp.width)).toBe(612);
    expect(Math.round(vp.height)).toBe(792);

    const scaled = page.getViewport({ scale: 1.5 });
    expect(Math.round(scaled.width)).toBe(918);

    const tc = await page.getTextContent();
    expect(tc.items.map((i) => ("str" in i ? i.str : "")).join(" ")).toContain(
      "Fixture page 2",
    );

    await task.destroy();
  });

  it("streams text content (streamTextContent used by the text layer)", async () => {
    const bytes = await makePdfBytes({ pages: 1 });
    const task = getDocument({
      data: bytes,
      standardFontDataUrl: STANDARD_FONT_DATA_URL,
    });
    const pdf = await task.promise;
    const page = await pdf.getPage(1);
    const reader = page.streamTextContent().getReader();
    let text = "";
    for (;;) {
      const r = await reader.read();
      if (r.done) break;
      for (const item of r.value.items) {
        if ("str" in item) text += item.str;
      }
    }
    expect(text).toContain("Fixture page 1");
    await task.destroy();
  });

  it("rejects a truncated PDF (magic intact, body cut) — error shape for CORRUPT mapping", async () => {
    const bytes = await makePdfBytes({ pages: 2 });
    const task = getDocument({
      data: truncated(bytes, 200),
      standardFontDataUrl: STANDARD_FONT_DATA_URL,
    });
    await expect(task.promise).rejects.toSatisfy((err: unknown) => {
      // 上游归 InvalidPDFException（render.test 的 CORRUPT 映射依据）
      return err instanceof Error && /InvalidPDF/i.test(err.name);
    });
  });

  it("rejects a valid-magic garbage body", async () => {
    const garbage = new Uint8Array(600);
    garbage.set([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x34], 0);
    for (let i = 8; i < garbage.length; i++) garbage[i] = 0x00;
    const task = getDocument({
      data: garbage,
      standardFontDataUrl: STANDARD_FONT_DATA_URL,
    });
    await expect(task.promise).rejects.toBeInstanceOf(Error);
  });
});
