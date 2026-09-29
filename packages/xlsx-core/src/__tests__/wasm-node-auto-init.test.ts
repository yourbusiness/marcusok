import { beforeEach, describe, expect, it, vi } from "vitest";

// Node auto-init unit tests: modern-xlsx is mocked WITH initWasmSync so the
// synchronous path is drivable, and node:fs is mocked so these tests never
// touch the real 2MB binary (the real end-to-end init lives in
// wasm-node-auto-init.integration.test.ts). The wasm URL candidates are real
// URLs derived from import.meta.url, so the path assertions below exercise
// the actual resolution shape.
const { initWasmMock, initWasmSyncMock, readFileSyncMock } = vi.hoisted(() => ({
  initWasmMock: vi.fn(),
  initWasmSyncMock: vi.fn(),
  readFileSyncMock: vi.fn(),
}));

vi.mock("modern-xlsx", () => ({
  initWasm: initWasmMock,
  initWasmSync: initWasmSyncMock,
}));
vi.mock("node:fs", () => ({
  readFileSync: readFileSyncMock,
}));

import { WasmLoader } from "../wasm-loader";

describe("WasmLoader Node auto-init (initWasmSync path)", () => {
  beforeEach(() => {
    initWasmMock.mockReset();
    initWasmSyncMock.mockReset();
    readFileSyncMock.mockReset();
    readFileSyncMock.mockReturnValue(new Uint8Array([1, 2, 3]));
  });

  it("initializes synchronously from the package's own wasm when no wasmUrl is configured", async () => {
    const loader = new WasmLoader();

    await loader.ensureLoaded();

    expect(loader.isReady).toBe(true);
    // The sync path ran exactly once; the async fetch path never did.
    expect(initWasmSyncMock).toHaveBeenCalledTimes(1);
    expect(initWasmMock).not.toHaveBeenCalled();
    // initWasmSync received what readFileSync returned...
    expect(initWasmSyncMock).toHaveBeenCalledWith(new Uint8Array([1, 2, 3]));
    // ...for the wasm binary resolved next to this module (a file URL ending
    // in modern-xlsx.wasm — the dist-next-to-entry candidate).
    const wasmPath = String(readFileSyncMock.mock.calls[0][0]);
    expect(wasmPath).toContain("modern-xlsx");
    expect(wasmPath).toMatch(/modern-xlsx\.wasm$/);
  });

  it("falls back to ../dist when the entry-adjacent wasm is missing (src mode)", async () => {
    // Running from src/ (repo tests, source-aliased monorepo consumers),
    // ./modern-xlsx.wasm does not exist and the ../dist candidate must take
    // over — exactly one retry, no initWasm degradation.
    readFileSyncMock.mockImplementationOnce(() => {
      throw new Error("ENOENT: no such file or directory");
    });
    const loader = new WasmLoader();

    await loader.ensureLoaded();

    expect(loader.isReady).toBe(true);
    expect(initWasmSyncMock).toHaveBeenCalledTimes(1);
    expect(initWasmMock).not.toHaveBeenCalled();
    expect(readFileSyncMock).toHaveBeenCalledTimes(2);
    const secondPath = String(readFileSyncMock.mock.calls[1][0]);
    expect(secondPath).toMatch(/[/\\]dist[/\\]modern-xlsx\.wasm$/);
  });

  it("degrades to the initWasm path when initWasmSync throws", async () => {
    // E.g. a bundled-for-Node build where the runtime package is not on disk.
    initWasmSyncMock.mockImplementation(() => {
      throw new Error("ENOENT: wasm binary missing");
    });
    initWasmMock.mockResolvedValue(undefined);
    const loader = new WasmLoader();

    await loader.ensureLoaded();

    expect(loader.isReady).toBe(true);
    expect(initWasmSyncMock).toHaveBeenCalledTimes(1);
    // No retry loop around the sync path: exactly one attempt, then the
    // standard initWasm run with the default URL.
    expect(initWasmMock).toHaveBeenCalledTimes(1);
    expect(String(initWasmMock.mock.calls[0][0])).toMatch(/modern-xlsx\.wasm$/);
  });

  it("never auto-inits when a wasmUrl is configured", async () => {
    initWasmMock.mockResolvedValue(undefined);
    const loader = new WasmLoader({ wasmUrl: "https://cdn.example/x.wasm" });

    await loader.ensureLoaded();

    expect(loader.isReady).toBe(true);
    expect(initWasmSyncMock).not.toHaveBeenCalled();
    expect(readFileSyncMock).not.toHaveBeenCalled();
    expect(initWasmMock).toHaveBeenCalledWith("https://cdn.example/x.wasm");
  });

  it("never auto-inits outside Node (browser-like global scope)", async () => {
    // process is the environment gate; without it the loader must take the
    // fetch path even with no URL configured.
    vi.stubGlobal("process", undefined);
    try {
      initWasmMock.mockResolvedValue(undefined);
      const loader = new WasmLoader();

      await loader.ensureLoaded();

      expect(loader.isReady).toBe(true);
      expect(initWasmSyncMock).not.toHaveBeenCalled();
      // No configured URL -> the default (shipped binary next to the entry).
      expect(String(initWasmMock.mock.calls[0][0])).toMatch(
        /modern-xlsx\.wasm$/,
      );
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
