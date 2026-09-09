import { beforeEach, describe, expect, it, vi } from "vitest";

// Node auto-init unit tests: modern-xlsx is mocked WITH initWasmSync so the
// synchronous path is drivable, and node:fs is mocked so these tests never
// touch the real 2MB binary (the real end-to-end init lives in
// wasm-node-auto-init.integration.test.ts). createRequire/node:path stay real,
// so the wasm path assertion below exercises the actual on-disk resolution.
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

  it("initializes synchronously from node_modules when no wasmUrl is configured", async () => {
    const loader = new WasmLoader();

    await loader.ensureLoaded();

    expect(loader.isReady).toBe(true);
    // The sync path ran exactly once; the async fetch path never did.
    expect(initWasmSyncMock).toHaveBeenCalledTimes(1);
    expect(initWasmMock).not.toHaveBeenCalled();
    // initWasmSync received what readFileSync returned...
    expect(initWasmSyncMock).toHaveBeenCalledWith(new Uint8Array([1, 2, 3]));
    // ...for the wasm binary resolved next to the real modern-xlsx entry.
    const wasmPath = readFileSyncMock.mock.calls[0][0];
    expect(wasmPath).toContain("modern-xlsx");
    expect(wasmPath).toMatch(/modern-xlsx\.wasm$/);
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
    // standard initWasm(undefined) run.
    expect(initWasmMock).toHaveBeenCalledTimes(1);
    expect(initWasmMock).toHaveBeenCalledWith(undefined);
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
      expect(initWasmMock).toHaveBeenCalledWith(undefined);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
