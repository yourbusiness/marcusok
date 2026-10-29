import { describe, expect, it } from "vitest";
import { sniffPpt } from "../sniff";
import { buildPptx, docxBytes, fakeOle2Bytes, plainZipBytes } from "./fixtures";

describe("sniffPpt", () => {
  it("空输入归 UNSUPPORTED", () => {
    const r = sniffPpt(new Uint8Array(0));
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error.code).toBe("UNSUPPORTED");
      expect(r.error.message).toContain("empty");
    }
  });

  it("合法 pptx（PK 魔数）放行", () => {
    expect(sniffPpt(buildPptx({ texts: ["a"] })).ok).toBe(true);
  });

  it("非 pptx 的 zip（PK 魔数）也放行——部件校验是引擎的职责", () => {
    expect(sniffPpt(plainZipBytes()).ok).toBe(true);
    expect(sniffPpt(docxBytes()).ok).toBe(true);
  });

  it("OLE2 魔数放行且不预判 legacy——加密/老式/未知 CFB 只能由引擎按目录区分", () => {
    const r = sniffPpt(fakeOle2Bytes());
    expect(r.ok).toBe(true); // 关键：不得在前置阶段报 LEGACY_FORMAT
  });

  it("截断的 zip（仍有 PK 头）放行，交给引擎报结构错误", () => {
    const bytes = buildPptx({ texts: ["x"] }).slice(0, 30);
    expect(sniffPpt(bytes).ok).toBe(true);
  });

  it("非 ZIP 非 OLE2 输入归 UNSUPPORTED", () => {
    for (const bytes of [
      new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]),
      new Uint8Array(1024).fill(0x41), // "AAAA..."
      new TextEncoder().encode("GIF89a....."),
    ]) {
      const r = sniffPpt(bytes);
      expect(r.ok).toBe(false);
      if (!r.ok) expect(r.error.code).toBe("UNSUPPORTED");
    }
  });

  it("PK 只匹配前缀——EOCD-only 空 zip（PK\x05\x06）同样放行", () => {
    const r = sniffPpt(new Uint8Array([0x50, 0x4b, 0x05, 0x06, 0, 0]));
    expect(r.ok).toBe(true);
  });
});
