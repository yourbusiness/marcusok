import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // 默认 node（sniff 纯字节测试）；渲染类测试在文件头用
    // `// @vitest-environment happy-dom` 注释切换环境（render/create/
    // security/multi-instance/performance），与 excel-preview 同一模式。
    environment: "node",
    include: ["src/__tests__/**/*.test.ts"],
    testTimeout: 60_000,
    hookTimeout: 60_000,
    // 性能基线并存：串行保基线稳定（excel-preview 同款结论）
    fileParallelism: false,
  },
});
