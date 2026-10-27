import { defineConfig } from "vitest/config";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const pkgRoot = dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);

// pdfjs-dist 的现代构建在 Node 下 getDocument 崩（worker 代码路径依赖
// 浏览器环境，实测报 hashOriginal.toHex is not a function，上游警告明示
// 用 legacy 构建）；Node 测试定向 alias 到 legacy 构建（同代码 + polyfill，
// parse/错误归一逻辑两端 API 同源）。浏览器发布走现代构建（tsup 主入口），
// 两侧差异由浏览器验收兜底（docs/play demo）。属本包自有测试基础设施，
// 不影响其他包。
const legacyBuild = resolve(
  dirname(require.resolve("pdfjs-dist/package.json")),
  "legacy",
  "build",
  "pdf.mjs",
);

export default defineConfig({
  resolve: {
    alias: [
      // 精确匹配裸导入（正则锚定，避免前缀误伤未来可能出现的子路径导入）
      { find: /^pdfjs-dist$/, replacement: legacyBuild },
    ],
  },
  test: {
    // 默认 node（sniff/parse 纯逻辑）；渲染类测试在文件头用
    // `// @vitest-environment happy-dom` 切换环境（word-preview 同一模式）。
    environment: "node",
    include: ["src/__tests__/**/*.test.ts"],
    testTimeout: 60_000,
    hookTimeout: 60_000,
    // 性能基线并存：串行保基线稳定（仓库三包同款结论）
    fileParallelism: false,
  },
});
