/**
 * sniff 纯字节测试（node 环境，不依赖 DOM）。
 *
 * 覆盖（计划第五节 P0 要求的五类坏包 + 加密位）：
 * 非zip、四文件 docx（正例）、zip 非docx、缺 word/document.xml、缺
 * _rels/.rels、截断 zip、空文件、OLE2、GP bit 0（加密位）。
 * 加密位检测是 excel-preview 没有的新代码路径（其加密判别走引擎错误码），
 * 必须专门覆盖。
 */
import { describe, expect, it } from "vitest";
import { sniffWordDocument } from "../sniff";
import {
  simpleDocx,
  docxMissingPart,
  encryptedZipBytes,
  fakeOle2Bytes,
  plainZipBytes,
} from "./fixtures";

describe("sniffWordDocument", () => {
  it("完整四文件 docx：放行", async () => {
    const r = await sniffWordDocument(simpleDocx("hello"));
    expect(r.ok).toBe(true);
  });

  it("空文件：UNSUPPORTED", async () => {
    const r = await sniffWordDocument(new Uint8Array(0));
    expect(r).toEqual({
      ok: false,
      error: { code: "UNSUPPORTED", message: expect.stringContaining("empty") },
    });
  });

  it("非 zip（纯文本）：UNSUPPORTED", async () => {
    const r = await sniffWordDocument(
      new TextEncoder().encode("plain text, not a zip"),
    );
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("UNSUPPORTED");
  });

  it("OLE2 魔数（.doc 老格式 / 加密 docx 同魔数）：LEGACY_FORMAT，message 合并表述", async () => {
    const r = await sniffWordDocument(fakeOle2Bytes());
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error.code).toBe("LEGACY_FORMAT");
      expect(r.error.message).toMatch(/legacy/i);
      expect(r.error.message).toMatch(/encrypted/i);
    }
  });

  it("zip 但非 docx（无 OPC 部件）：UNSUPPORTED 而非上游模糊空引用错", async () => {
    const r = await sniffWordDocument(plainZipBytes());
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error.code).toBe("UNSUPPORTED");
      // message 指名缺失部件，可读
      expect(r.error.message).toContain("word/document.xml");
    }
  });

  it("缺 word/document.xml：UNSUPPORTED", async () => {
    const r = await sniffWordDocument(docxMissingPart("word/document.xml"));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("UNSUPPORTED");
  });

  it("缺 _rels/.rels：UNSUPPORTED", async () => {
    const r = await sniffWordDocument(docxMissingPart("_rels/.rels"));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("UNSUPPORTED");
  });

  it("缺 [Content_Types].xml：UNSUPPORTED", async () => {
    const r = await sniffWordDocument(docxMissingPart("[Content_Types].xml"));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("UNSUPPORTED");
  });

  it("截断 zip（中央目录不可达）：CORRUPT（zip 结构坏，而非『不是 Word』）", async () => {
    const bytes = simpleDocx("hello").slice(0, 40);
    const r = await sniffWordDocument(bytes);
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error.code).toBe("CORRUPT");
      expect(r.error.cause).toBeDefined();
    }
  });

  it("zip general purpose bit 0（加密位）：PASSWORD_PROTECTED", async () => {
    const r = await sniffWordDocument(encryptedZipBytes());
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error.code).toBe("PASSWORD_PROTECTED");
      expect(r.error.message).toMatch(/password/i);
    }
  });

  it("OLE2 魔数优先于 zip 检查（OLE2 头不含 PK，天然分流）", async () => {
    // 512 字节 OLE2 块中后续字节可能有任意值，魔数检查必须先行命中
    const r = await sniffWordDocument(fakeOle2Bytes());
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).not.toBe("UNSUPPORTED");
  });
});
