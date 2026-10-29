/**
 * Post-build step: ship the parser wasm in this package's dist.
 *
 * `pptx_parser_bg.wasm`（~1.95MB）是 main 模式解析的硬需求：上游 chunk 以
 * `new URL("pptx_parser_bg.wasm", import.meta.url)` 定位它，bundler 消费方
 * （Vite 等）只有在包 dist 内真实存在该文件时才会 emit + 重写 URL，否则
 * 运行时按 import.meta.url 解析到不存在的路径直接 404（spike 实测，见
 * docs/ppt-preview-spike-record.md S1）。
 *
 * 上游的 worker 资产（pptx-source-worker.mjs、assets/render-worker-*.js）
 * 有意不拷贝：它们只服务 mode:'worker' 渲染模式与 legacy 路径，本包固定
 * main 模式不触达；不拷贝会在消费方构建日志留下 2 条可忽略警告。
 *
 * 定位方式：require.resolve("@silurus/ooxml/pptx") 命中 dist/pptx.mjs，
 * 再向上推包根——上游 exports 不暴露 ./package.json，直接 resolve 会抛
 * ERR_PACKAGE_PATH_NOT_EXPORTED（spike 实测），禁止改成那种写法。
 *
 * Runs via the main config's onSuccess hook so watch rebuilds re-copy too
 * (clean:true wipes dist on watch start — excel-exporter 2.1.5 lesson).
 */
import { copyFileSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const distDir = resolve(dirname(fileURLToPath(import.meta.url)), "..", "dist");

// @silurus/ooxml 是 devDependency（构建期打进 dist）；经 Node 解析定位安装
// 目录，pnpm 的符号链接布局对其不透明。
const pptxEntry = require.resolve("@silurus/ooxml/pptx");
const pkgRoot = dirname(pptxEntry); // .../node_modules/@silurus/ooxml/dist

const wasmSrc = resolve(pkgRoot, "pptx_parser_bg.wasm");
if (!statSync(wasmSrc, { throwIfNoEntry: false })) {
  throw new Error(
    `[ppt-preview] pptx_parser_bg.wasm not found at ${wasmSrc}. ` +
      "Run pnpm install first (@silurus/ooxml is a devDependency of this package).",
  );
}

copyFileSync(wasmSrc, resolve(distDir, "pptx_parser_bg.wasm"));
console.log(`[ppt-preview] forwarded pptx_parser_bg.wasm -> ${distDir}`);
