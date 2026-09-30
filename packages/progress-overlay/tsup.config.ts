import { defineConfig } from "tsup";

// 单入口零依赖：遮罩是纯 DOM + CSS 字符串 + 状态机，无任何运行时依赖。
// 本包是仓库私有层（不再发布 npm）：excel-exporter 以构建期依赖引用，
// 其 tsup 把本包 dist 整体打进导出包产物——外部唯一取用点是
// @marcusok/excel-exporter/overlay 子路径（口径见 src/index.ts 头注释
// 与包 README）。
export default defineConfig({
  entry: { index: "src/index.ts" },
  format: ["esm"],
  dts: true,
  splitting: false,
  treeshake: true,
  clean: true,
  sourcemap: true,
  target: "es2022",
});
