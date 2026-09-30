import { defineConfig } from "tsup";

// 单入口零依赖：遮罩是纯 DOM + CSS 字符串 + 状态机，不打进任何运行时
// 依赖，消费方（excel-exporter 及未来的导出/预览包）以 external 方式引用，
// 同页面多包共享一份实现与样式注入。
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
