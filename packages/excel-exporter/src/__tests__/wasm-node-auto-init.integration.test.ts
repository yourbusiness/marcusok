import { describe, expect, it } from "vitest";

// No mocks anywhere in this file: it exercises the REAL zero-configuration
// Node path a published-package consumer gets — no configureWasm call, no
// initWasmSync boilerplate. The loader must locate this package's shipped
// wasm on disk (import.meta.url-relative, symlink-safe), init it
// synchronously, and a small export must come back styled-engine
// ("modern-xlsx" mode "main", not the style-less stream fallback).
import { exportExcel, getWasmLoader } from "../index";
import { WasmLoader, defaultWasmUrl } from "@marcusok/xlsx-core";

describe("Node auto-init (integration, real wasm binary)", () => {
  it("loads the real modern-xlsx.wasm from node_modules with zero configuration", async () => {
    const loader = getWasmLoader();

    await loader.ensureLoaded();

    expect(loader.isReady).toBe(true);
  });

  it("accepts an explicitly configured file: wasmUrl in Node", async () => {
    // Self-hosted / Electron deployments pass a local path; Node's fetch
    // rejects file: URLs, so the loader must read it synchronously.
    const fresh = new WasmLoader({ wasmUrl: defaultWasmUrl() });
    await fresh.ensureLoaded();
    expect(fresh.isReady).toBe(true);
  });

  it("reports the real fs reason for a missing configured file", async () => {
    const fresh = new WasmLoader({
      wasmUrl: new URL("./does-not-exist.wasm", import.meta.url),
    });
    await expect(fresh.ensureLoaded()).rejects.toThrow(
      /Failed to load the configured WASM file .*does-not-exist\.wasm/,
    );
  });

  it("exports through the modern-xlsx engine without any configuration", async () => {
    const result = await exportExcel({
      filename: "auto-init-check",
      download: false,
      sheets: [
        {
          name: "S",
          columns: [{ prop: "a", label: "A" }],
          data: [{ a: 1 }, { a: 2 }],
        },
      ],
    });

    expect(result.success).toBe(true);
    // engine proves the styled WASM engine ran; a degraded export would come
    // back mode "stream" with a styles-stripped soft error.
    expect(result.engine).toBe("modern-xlsx");
    expect(result.mode).toBe("main");
    expect(result.blob).toBeInstanceOf(Blob);
    expect(result.error).toBeUndefined();
  });
});
