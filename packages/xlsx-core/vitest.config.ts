import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["src/__tests__/**/*.test.ts"],
    testTimeout: 60_000,
    hookTimeout: 60_000,
    // The integration test compiles the real 2MB wasm binary; keeping files
    // serial avoids CPU spikes landing next to each other (pattern carried
    // over from excel-exporter, where it also protects its perf baselines).
    fileParallelism: false,
  },
});
