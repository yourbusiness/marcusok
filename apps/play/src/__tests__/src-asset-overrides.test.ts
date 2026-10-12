import { describe, expect, it } from "vitest";
import { rewriteAssetLiterals } from "../vite/src-asset-overrides";

describe("rewriteAssetLiterals（字面量 → ?url 导入注入）", () => {
  it("命中字面量：替换为注入变量并顶部加 ?url 导入", () => {
    const code =
      'const url = parseWorkerUrl ?? new URL("./parse.worker.js", import.meta.url);';
    const out = rewriteAssetLiterals(code, "@marcusok/excel-preview", [
      "parse.worker.js",
    ])!;
    expect(out).toContain(
      'import __marcusok_asset__marcusok_excel_preview_0 from "@marcusok/excel-preview/dist/parse.worker.js?url";',
    );
    expect(out).toContain("const url = parseWorkerUrl ?? __marcusok_asset_");
    // 注入的 import 必须在模块顶部（transform 产物需是合法 ESM）
    expect(out.indexOf("import ")).toBe(0);
  });

  it("多处出现同一字面量全部替换（同名导入只注入一次）", () => {
    const code = [
      'const a = new URL("./modern-xlsx.wasm", import.meta.url);',
      'const b = new URL("./modern-xlsx.wasm", import.meta.url);',
    ].join("\n");
    const out = rewriteAssetLiterals(code, "@marcusok/xlsx-core", [
      "modern-xlsx.wasm",
    ])!;
    expect(out.match(/import __marcusok/g)?.length).toBe(1);
    expect(out).not.toContain('new URL("./modern-xlsx.wasm"');
    expect(out).toContain("const a = __marcusok_asset_");
    expect(out).toContain("const b = __marcusok_asset_");
  });

  it("无命中返回 null（插件协议：不改写）", () => {
    expect(
      rewriteAssetLiterals("const a = 1;", "@marcusok/xlsx-core", [
        "parse.worker.js",
      ]),
    ).toBeNull();
  });

  it("多文件各自注入独立变量（下标区分）", () => {
    const code = [
      'new URL("./a.js", import.meta.url)',
      'new URL("./b.wasm", import.meta.url)',
    ].join("\n");
    const out = rewriteAssetLiterals(code, "@marcusok/x", ["a.js", "b.wasm"])!;
    expect(out).toContain('from "@marcusok/x/dist/a.js?url"');
    expect(out).toContain('from "@marcusok/x/dist/b.wasm?url"');
  });
});
