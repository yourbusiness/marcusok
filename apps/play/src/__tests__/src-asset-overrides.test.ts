import { describe, expect, it } from "vitest";
import { rewriteAssetLiterals } from "../vite/src-asset-overrides";

describe("rewriteAssetLiterals（字面量 → ?url 导入注入）", () => {
  it("命中字面量：替换为 new URL(注入变量) 并顶部加 ?url 导入", () => {
    const code =
      'const url = parseWorkerUrl ?? new URL("./parse.worker.js", import.meta.url);';
    const out = rewriteAssetLiterals(code, "@marcusok/excel-preview", [
      "parse.worker.js",
    ])!;
    expect(out).toContain(
      'import __marcusok_asset__marcusok_excel_preview_0 from "@marcusok/excel-preview/dist/parse.worker.js?url";',
    );
    expect(out).toContain(
      "const url = parseWorkerUrl ?? new URL(__marcusok_asset__marcusok_excel_preview_0, import.meta.url)",
    );
    // 注入的 import 必须在模块顶部（transform 产物需是合法 ESM）
    expect(out.indexOf("import ")).toBe(0);
  });

  it("替换结果保持 URL 对象语义：.href 等成员访问不被破坏", () => {
    // pdf-preview 的 workerSrc 形态——曾因裸字符串替换让 .href 变 undefined，
    // worker 在 play 的 dev/build 双模式静默失效（"Invalid workerSrc"）
    const code =
      'return new URL("./pdf.worker.min.mjs", import.meta.url).href;';
    const out = rewriteAssetLiterals(code, "@marcusok/pdf-preview", [
      "pdf.worker.min.mjs",
    ])!;
    expect(out).toContain(
      "return new URL(__marcusok_asset__marcusok_pdf_preview_0, import.meta.url).href;",
    );
    // 裸标识符形态（丢 URL 对象语义）必须不存在
    expect(out).not.toMatch(/return __marcusok_asset_[\s\S]*?\.href/);
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
    expect(out).toContain("const a = new URL(__marcusok_asset_");
    expect(out).toContain("const b = new URL(__marcusok_asset_");
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
