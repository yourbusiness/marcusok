/**
 * resolveWasmUrl 分支覆盖：资产链路的核心契约（显式透传 / 空串拒绝 /
 * src 锚点 ../dist 兜底 / bundler emit 场景不干预）此前无回归保护。
 *
 * 锚点说明：vitest 直跑 src，默认锚点（import.meta.url）天然落在 /src/，
 * 走兜底分支；bundler emit 分支（非 src 锚点 → undefined）通过注入
 * anchorBase 覆盖。
 */
import { describe, expect, it } from "vitest";

import { resolveWasmUrl } from "../assets";

describe("resolveWasmUrl", () => {
  it("URL 实例直接透传（同一实例）", () => {
    const url = new URL("https://cdn.example.com/asset/pptx_parser_bg.wasm");
    expect(resolveWasmUrl(url)).toBe(url);
  });

  it("非空字符串透传", () => {
    expect(resolveWasmUrl("/self-hosted/pptx_parser_bg.wasm")).toBe(
      "/self-hosted/pptx_parser_bg.wasm",
    );
  });

  it("空字符串抛 TypeError（配置错误 fail-fast）", () => {
    expect(() => resolveWasmUrl("")).toThrow(TypeError);
  });

  it("不传 + 锚点落在 src/（play src 联调）→ ../dist/pptx_parser_bg.wasm 兜底", () => {
    const url = resolveWasmUrl(
      undefined,
      "file:///repo/packages/ppt-preview/src/index.ts",
    );
    expect(url).toBeInstanceOf(URL);
    expect((url as URL).pathname).toBe(
      "/repo/packages/ppt-preview/dist/pptx_parser_bg.wasm",
    );
  });

  it("不传 + 默认锚点（vitest 直跑 src）→ 同样命中 ../dist 兜底", () => {
    const url = resolveWasmUrl(undefined);
    expect(url).toBeInstanceOf(URL);
    expect((url as URL).pathname).toMatch(/\/dist\/pptx_parser_bg\.wasm$/);
  });

  it("不传 + 锚点在 dist/（bundler emit 场景）→ undefined，不干预上游默认", () => {
    expect(
      resolveWasmUrl(undefined, "https://cdn.example.com/pkg/dist/index.js"),
    ).toBeUndefined();
  });
});
