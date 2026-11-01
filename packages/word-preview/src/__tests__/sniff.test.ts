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
import { strToU8, unzipSync, zipSync } from "fflate";
import { sniffWordDocument } from "../sniff";
import {
  simpleDocx,
  centralDirEncryptedZipBytes,
  docxMissingPart,
  docxWithMainPart,
  encryptedZipBytes,
  fakeOle2Bytes,
  plainZipBytes,
} from "./fixtures";

const WORD_MAIN_CT =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

const CT_XML = (mainPart: string, contentType: string): string =>
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/${mainPart}" ContentType="${contentType}"/></Types>`;

const ROOT_RELS_XML = (target: string): string =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${target}"/></Relationships>`;

/** 与 docx 同构的 OPC 包（部件齐全，仅主部件路径/类型不同）。 */
function opcBytes(mainPart: string, contentType: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(CT_XML(mainPart, contentType)),
    "_rels/.rels": strToU8(ROOT_RELS_XML(mainPart)),
    [mainPart]: strToU8("<root/>"),
  });
}

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

  it("只有中央目录头带加密位（local header 未标记）：同样 PASSWORD_PROTECTED", async () => {
    // JSZip 在 readCentralPart 抛 "Encrypted zip are not supported"；
    // 旧实现把它统一归 CORRUPT，与文档承诺的 PASSWORD_PROTECTED 不符。
    const r = await sniffWordDocument(centralDirEncryptedZipBytes());
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error.code).toBe("PASSWORD_PROTECTED");
      expect(r.error.message).toMatch(/password/i);
    }
  });

  it("以 PK 开头但不是 ZIP 签名（如 PKG 文本）：UNSUPPORTED", async () => {
    // 旧实现只校验前两字节 PK；"PKG archive" 被送进 JSZip，
    // 抛 "end of central directory" 后误归 CORRUPT。
    const r = await sniffWordDocument(
      new TextEncoder().encode("PKG archive, not a zip"),
    );
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("UNSUPPORTED");
  });

  it("主部件由根关系的 officeDocument Target 决定（非 word/document.xml）", async () => {
    // OPC 允许主部件使用任意路径；上游按 _rels/.rels 解析，sniff 不能
    // 硬编码 word/document.xml 把合法文档拒之门外。
    const r = await sniffWordDocument(docxWithMainPart("word/main.xml"));
    expect(r.ok).toBe(true);
  });

  it("主部件目标缺失时：UNSUPPORTED 且 message 指名实际路径", async () => {
    const bytes = docxWithMainPart("word/main.xml");
    // 用 zipSync 重建并删掉主部件
    const files: Record<string, Uint8Array> = {};
    for (const [k, v] of Object.entries(unzipSync(bytes))) {
      if (k !== "word/main.xml") files[k] = v;
    }
    const r = await sniffWordDocument(zipSync(files));
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error.code).toBe("UNSUPPORTED");
      expect(r.error.message).toContain("word/main.xml");
    }
  });

  it("OLE2 魔数优先于 zip 检查（OLE2 头不含 PK，天然分流）", async () => {
    // 512 字节 OLE2 块中后续字节可能有任意值，魔数检查必须先行命中
    const r = await sniffWordDocument(fakeOle2Bytes());
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).not.toBe("UNSUPPORTED");
  });

  it("同构 OPC 但主部件是 xlsx/pptx：UNSUPPORTED（不是损坏的 Word）", async () => {
    // 三种 OPC 包都有 [Content_Types].xml、_rels/.rels 与 officeDocument
    // 主部件；旧实现只校验部件存在，xlsx/pptx 会被送进 Word 引擎并以
    // CORRUPT + 引擎内部错误收场。主部件 ContentType 才是类型标识。
    for (const [mainPart, contentType] of [
      [
        "xl/workbook.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
      ],
      [
        "ppt/presentation.xml",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
      ],
    ] as const) {
      const r = await sniffWordDocument(opcBytes(mainPart, contentType));
      expect(r.ok, mainPart).toBe(false);
      if (!r.ok) {
        expect(r.error.code).toBe("UNSUPPORTED");
        expect(r.error.message).toContain("Not a Word document");
        expect(r.error.message).toContain(contentType);
      }
    }
  });

  it("docm（macroEnabled 主部件）放行", async () => {
    const r = await sniffWordDocument(
      opcBytes(
        "word/document.xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
      ),
    );
    expect(r.ok).toBe(true);
  });

  it("关系 Target 使用十六进制字符引用时仍能定位主部件", async () => {
    // 与引擎的 DOMParser 解码口径对齐：&#x64; = "d"。
    const bytes = zipSync({
      "[Content_Types].xml": strToU8(CT_XML("word/document.xml", WORD_MAIN_CT)),
      "_rels/.rels": strToU8(ROOT_RELS_XML("word/&#x64;ocument.xml")),
      "word/document.xml": strToU8("<w:document/>"),
      "word/_rels/document.xml.rels": strToU8("<Relationships/>"),
    });
    const r = await sniffWordDocument(bytes);
    expect(r.ok).toBe(true);
  });
});
