import { defineConfig } from "tsup";

// 单配置：docx-preview 连同其唯一 runtime 依赖 jszip 打进本包 dist
// （发布面收敛为业务包，两者全在 devDependencies——npm 消费方零传递依赖）。
// devDeps 默认本就被 tsup 打包，显式 noExternal 是意图文档 + 防 tsup 行为
// 变化（与 excel-preview 的 worker 配置同风格）。
//
// platform 必须保持 browser：jszip 的 browser 字段映射到自包含的
// dist/jszip.min.js（lie/pako/readable-stream/setimmediate 四个 runtime
// 依赖已内联），改成 node 会使打包图突变、把传递依赖泄漏给消费方。
//
// 不 minify（与 excel-preview 一致）：esbuild 对 bundle 默认
// legalComments=eof，docx-preview（Apache-2.0）与 jszip（MIT）的许可头
// 注释会保留在产物尾部；构建后须 grep 验证（见 README 构建说明）。
export default defineConfig({
  entry: { index: "src/index.ts" },
  format: ["esm"],
  dts: { resolve: true },
  splitting: true,
  treeshake: true,
  clean: true,
  sourcemap: true,
  target: "es2022",
  noExternal: ["docx-preview", "jszip"],
  platform: "browser",
});
