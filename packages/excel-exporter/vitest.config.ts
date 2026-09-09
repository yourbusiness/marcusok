import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["src/__tests__/**/*.test.ts"],
    testTimeout: 60_000,
    hookTimeout: 60_000,
    // Several files compile the real 2MB wasm binary (setup.ts bootstrap, the
    // auto-init integration tests). Under default file parallelism those CPU
    // spikes land next to performance.test.ts timing cases and push them past
    // the SLA threshold (observed 10k case at 204-273ms vs ~120ms isolated).
    // Serial files keep the timing baselines meaningful; the suite still
    // finishes in a few seconds.
    fileParallelism: false,
  },
});
