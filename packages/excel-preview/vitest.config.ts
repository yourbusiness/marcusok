import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // 默认 node（模型/格式化）；渲染类测试在文件头用
    // `// @vitest-environment happy-dom` 注释切换环境（render/performance/
    // worker-domain 三个文件），无需路径映射表。
    environment: "node",
    include: ["src/__tests__/**/*.test.ts"],
    testTimeout: 60_000,
    hookTimeout: 60_000,
    // wasm 编译与性能基线并存：串行保基线稳定（excel-exporter 同款结论）
    fileParallelism: false,
  },
});
