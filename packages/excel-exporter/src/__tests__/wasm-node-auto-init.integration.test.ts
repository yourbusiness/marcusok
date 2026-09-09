import { describe, expect, it } from "vitest";

// No mocks anywhere in this file: it exercises the REAL zero-configuration
// Node path a published-package consumer gets — no configureWasm call, no
// initWasmSync boilerplate. The loader must locate modern-xlsx's wasm through
// node_modules (createRequire, pnpm-symlink-safe), init it synchronously, and
// a small export must come back styled-engine ("modern-xlsx", not the
// style-less SheetJS fallback).
import { exportExcel, getWasmLoader } from "../index";

describe("Node auto-init (integration, real wasm binary)", () => {
  it("loads the real modern-xlsx.wasm from node_modules with zero configuration", async () => {
    const loader = getWasmLoader();

    await loader.ensureLoaded();

    expect(loader.isReady).toBe(true);
  });

  it("exports through the modern-xlsx engine without any configuration", async () => {
    const result = await exportExcel({
      filename: "auto-init-check",
      download: false,
      sheets: [
        {
          name: "S",
          columns: [{ key: "a", header: "A" }],
          data: [{ a: 1 }, { a: 2 }],
        },
      ],
    });

    expect(result.success).toBe(true);
    // engine proves the styled WASM engine ran; a degraded export would be
    // "sheetjs" with a styles-stripped warning.
    expect(result.engine).toBe("modern-xlsx");
    expect(result.mode).toBe("main");
    expect(result.blob).toBeInstanceOf(Blob);
    expect(result.error).toBeUndefined();
  });
});
