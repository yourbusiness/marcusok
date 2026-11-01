// @vitest-environment happy-dom
/**
 * 资产 URL 推导回归测试。
 *
 * 关键回归点：显式 workerSrc 可以是根相对/相对路径（自托管最常见写法），
 * 推导 cmaps/standard_fonts 时不得把它直接当 URL base——`new URL("cmaps/",
 * "/assets/x.mjs")` 会抛 `Invalid base URL`，而该异常会被归成 CORRUPT，
 * 让用户误以为 PDF 损坏。
 */
import { describe, it, expect } from "vitest";
import { resolveAssetUrls } from "../assets";

describe("resolveAssetUrls", () => {
  it("keeps an absolute workerSrc and derives sibling asset directories", () => {
    const out = resolveAssetUrls({
      workerSrc: "https://cdn.example.com/pdf/worker.mjs",
    });
    expect(out).toEqual({
      workerSrc: "https://cdn.example.com/pdf/worker.mjs",
      cMapUrl: "https://cdn.example.com/pdf/cmaps/",
      standardFontDataUrl: "https://cdn.example.com/pdf/standard_fonts/",
    });
  });

  it("resolves a root-relative workerSrc against the page base", () => {
    const out = resolveAssetUrls({
      workerSrc: "/assets/pdf.worker.min.mjs",
    });
    expect(out.workerSrc).toBe("/assets/pdf.worker.min.mjs");
    expect(new URL(out.cMapUrl).pathname).toBe("/assets/cmaps/");
    expect(new URL(out.standardFontDataUrl).pathname).toBe(
      "/assets/standard_fonts/",
    );
  });

  it("resolves a relative workerSrc against the page base", () => {
    const out = resolveAssetUrls({ workerSrc: "./vendor/worker.mjs" });
    expect(new URL(out.cMapUrl).pathname).toBe("/vendor/cmaps/");
    expect(new URL(out.standardFontDataUrl).pathname).toBe(
      "/vendor/standard_fonts/",
    );
  });

  it("falls back to package assets for non-hierarchical workerSrc (blob/data)", () => {
    // blob:/data: 不能作为相对 URL 的 base；旧实现直接抛 Invalid URL，
    // 合法的 object-URL worker 无法使用默认 cmaps/fonts。
    for (const workerSrc of [
      "blob:https://example.com/8c1b3d2a-0000-4000-8000-000000000000",
      "data:text/javascript;base64,ZXhwb3J0IHt9Ow==",
    ]) {
      const out = resolveAssetUrls({ workerSrc });
      expect(out.workerSrc).toBe(workerSrc);
      expect(out.cMapUrl.endsWith("/cmaps/")).toBe(true);
      expect(out.standardFontDataUrl.endsWith("/standard_fonts/")).toBe(true);
    }
  });

  it("lets explicit cMapUrl/standardFontDataUrl win over derivation", () => {
    const out = resolveAssetUrls({
      workerSrc: "/assets/worker.mjs",
      cMapUrl: "/cm/",
      standardFontDataUrl: "/sf/",
    });
    expect(out.cMapUrl).toBe("/cm/");
    expect(out.standardFontDataUrl).toBe("/sf/");
  });
});
