import { defineConfig } from "tsup";

// 单配置：@silurus/ooxml 的 pptx 子路径整体打进本包 dist（word-preview 打
// docx-preview / pdf-preview 打 pdfjs-dist 的同款决策）。上游只在
// devDependencies——npm 消费方零传递依赖。
//
// noExternal 用正则匹配子路径：本包 import 的是 `@silurus/ooxml/pptx`
// （根入口只导出命名空间，实现必须走 /pptx，见 docs/ spike 记录），字符串
// 精确匹配无法覆盖子路径导入。
//
// platform browser：剔除上游的 Node 分支。上游 main 模式（本包固定）的解析
// Worker 在运行时由 Blob/data: URL 内联创建，默认路径不需要外部 worker 文件；
// 消费方 bundler 会看到两条指向 worker 资产（pptx-source-worker.mjs、
// assets/render-worker-*.js）的 new URL 构建期警告——那是 mode:'worker'/
// legacy 分支的引用，main 模式不触达，属可忽略噪音（limits 文档已说明）。
// 唯一必需的外部资产是 wasm，由 copy-assets.mjs 从上游 dist 原样拷入
// （bundler 检测到存在即 emit 并重写 URL；不拷贝则运行时 404——spike 实测）。
export default defineConfig({
  entry: { index: "src/index.ts" },
  format: ["esm"],
  dts: { resolve: true },
  splitting: true,
  treeshake: true,
  clean: true,
  sourcemap: true,
  target: "es2022",
  noExternal: [/@silurus\/ooxml.*/],
  onSuccess: "node scripts/copy-assets.mjs",
  platform: "browser",
});
