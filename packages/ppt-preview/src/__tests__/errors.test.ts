import { describe, expect, it } from "vitest";
import { normalizeError } from "../errors";

/** 构造带 code 的 Error（vi.mock/真实上游都是 Error 子类挂 code）。 */
function ooxml(code: string, message = `upstream ${code}`): Error {
  const e = new Error(message);
  (e as Error & { code: string }).code = code;
  return e;
}

describe("normalizeError 映射", () => {
  it.each([
    ["legacy-binary-format", "LEGACY_FORMAT"],
    ["unsupported-encryption", "UNSUPPORTED"],
    ["unsupported-content", "UNSUPPORTED"],
    ["unsupported-image-effect", "UNSUPPORTED"],
    ["invalid-numbering", "CORRUPT"],
  ] as const)("%s → %s", (code, expected) => {
    const r = normalizeError(ooxml(code));
    expect(r.code).toBe(expected);
    expect(r.cause).toBeDefined();
  });

  it("encrypted → PASSWORD_PROTECTED（缺密码口径）", () => {
    const r = normalizeError(
      ooxml("encrypted", "This file is password-protected"),
    );
    expect(r.code).toBe("PASSWORD_PROTECTED");
    expect(r.message).toContain("password");
  });

  it("invalid-password → PASSWORD_PROTECTED（密码错误口径，message 不同于缺密码）", () => {
    const r = normalizeError(
      ooxml("invalid-password", "The supplied password is incorrect."),
    );
    expect(r.code).toBe("PASSWORD_PROTECTED");
    expect(r.message).toContain("rejected");
    const missing = normalizeError(
      ooxml("encrypted", "This file is password-protected"),
    );
    expect(r.message).not.toBe(missing.message);
  });

  it("not-ooxml → UNSUPPORTED（无 stage 可分流，不解析 message 细分损坏原因）", () => {
    const r = normalizeError(
      ooxml("not-ooxml", "This file is not an Office Open XML document: ..."),
    );
    expect(r.code).toBe("UNSUPPORTED");
    expect(r.message).toContain("Not a .pptx");
  });

  it("OoxmlResourceLimitError（code ooxml-resource-limit）→ UNSUPPORTED，message 含预算口径", () => {
    const err = ooxml(
      "ooxml-resource-limit",
      "declared-inflated-bytes limit=104857600 observed=536870912",
    );
    const r = normalizeError(err);
    expect(r.code).toBe("UNSUPPORTED");
    expect(r.message).toContain("archive resource");
    expect(r.message).toContain("104857600");
  });

  it("OoxmlDecodedImageLimitError（RangeError, code ooxml-decoded-image-limit）→ UNSUPPORTED", () => {
    const err = new RangeError("image-dimension limit=8192 observed=16384");
    (err as RangeError & { code: string }).code = "ooxml-decoded-image-limit";
    const r = normalizeError(err);
    expect(r.code).toBe("UNSUPPORTED");
    expect(r.message).toContain("decoded image");
  });

  it("资源限制识别兼容无 code 的 name 兜底（vi.mock 场景 instanceof 失效）", () => {
    const e1 = new Error("limit exceeded");
    e1.name = "OoxmlResourceLimitError";
    expect(normalizeError(e1).code).toBe("UNSUPPORTED");
    const e2 = new RangeError("image too large");
    e2.name = "OoxmlDecodedImageLimitError";
    expect(normalizeError(e2).code).toBe("UNSUPPORTED");
  });

  it.each([
    ["Worker is not defined", undefined],
    ["Failed to fetch dynamically imported module", undefined],
    ["wasm streaming compile failed: TypeError", undefined],
    ["upstream 新增的错误码", "future-code"],
  ] as const)(
    "Worker/WASM/未知异常 → UNKNOWN，不误报文件损坏（%s）",
    (message, code) => {
      const err = new Error(message);
      if (code) (err as Error & { code: string }).code = code;
      const r = normalizeError(err);
      expect(r.code).toBe("UNKNOWN");
    },
  );

  it("非 Error 输入（字符串/undefined）→ UNKNOWN 且 message 可读", () => {
    expect(normalizeError("boom").code).toBe("UNKNOWN");
    expect(normalizeError(undefined).code).toBe("UNKNOWN");
  });
});
