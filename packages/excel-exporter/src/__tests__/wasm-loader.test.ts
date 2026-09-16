import { beforeEach, describe, expect, it, vi } from "vitest";

// modern-xlsx is fully mocked: the loader under test must be drivable into
// its error state on demand (the real initWasm is idempotent, so a previously
// initialized module can never be made to fail again).
const { initWasmMock } = vi.hoisted(() => ({
  initWasmMock: vi.fn(),
}));

vi.mock("modern-xlsx", () => ({
  initWasm: initWasmMock,
}));

import { WasmLoader } from "../wasm-loader";

function makeLoader(wasmUrl?: string): WasmLoader {
  return new WasmLoader({ wasmUrl, timeoutMs: 50, maxRetries: 2 });
}

describe("WasmLoader error recovery", () => {
  beforeEach(() => {
    initWasmMock.mockReset();
  });

  it("parks in the error state after retries are exhausted and refuses to re-run", async () => {
    initWasmMock.mockRejectedValue(new Error("boom"));
    const loader = makeLoader("bad.wasm");

    await expect(loader.ensureLoaded()).rejects.toThrow(
      /failed after 2 attempts: boom/,
    );
    expect(initWasmMock).toHaveBeenCalledTimes(2);
    expect(loader.isReady).toBe(false);

    // Without a reconfiguration the loader must not silently retry.
    await expect(loader.ensureLoaded()).rejects.toThrow(/previously failed/);
    expect(initWasmMock).toHaveBeenCalledTimes(2);
  });

  it("clears the error state on ANY updateOptions call, so the next load retries", async () => {
    initWasmMock.mockRejectedValue(new Error("network down"));
    const loader = makeLoader("good.wasm");
    await expect(loader.ensureLoaded()).rejects.toThrow(/failed after 2/);

    // Same URL, only a timeout tweak: previously this left the loader stuck
    // in "error" forever despite the error message suggesting a retry.
    loader.updateOptions({ timeoutMs: 60 });
    initWasmMock.mockResolvedValue(undefined);

    await loader.ensureLoaded();
    expect(loader.isReady).toBe(true);
    // Two failed attempts + one successful retry.
    expect(initWasmMock).toHaveBeenCalledTimes(3);
  });

  it("re-attempts initWasm with the new URL when wasmUrl changes on a ready loader", async () => {
    // Loader state-machine contract only: updateOptions resets a ready loader
    // so the next ensureLoaded re-CALLS initWasm with the new URL. Note the
    // real modern-xlsx initWasm is idempotent ("first successful init wins"),
    // so on a thread that already initialized WASM this second call is a
    // no-op — updateOptions warns about exactly that (next test).
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      initWasmMock.mockResolvedValue(undefined);
      const loader = makeLoader("a.wasm");
      await loader.ensureLoaded();
      expect(loader.isReady).toBe(true);
      expect(initWasmMock).toHaveBeenCalledTimes(1);

      loader.updateOptions({ wasmUrl: "b.wasm" });
      expect(loader.isReady).toBe(false);

      await loader.ensureLoaded();
      expect(loader.isReady).toBe(true);
      expect(initWasmMock).toHaveBeenCalledTimes(2);
      expect(initWasmMock).toHaveBeenLastCalledWith("b.wasm");
    } finally {
      warn.mockRestore();
    }
  });

  it("warns that a wasmUrl change cannot reload an already-initialized module", async () => {
    // Real-dependency caveat surfaced at the API boundary: modern-xlsx's
    // initWasm keeps the first successfully loaded module, so the reset+retry
    // above is silently a no-op on this thread. updateOptions must say so.
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      initWasmMock.mockResolvedValue(undefined);
      const loader = makeLoader("a.wasm");
      await loader.ensureLoaded();
      expect(warn).not.toHaveBeenCalled();

      loader.updateOptions({ wasmUrl: "b.wasm" });
      expect(warn).toHaveBeenCalledTimes(1);
      expect(String(warn.mock.calls[0][0])).toContain("idempotent");
    } finally {
      warn.mockRestore();
    }
  });

  it("keeps a ready loader ready when only timeouts/retries change", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      initWasmMock.mockResolvedValue(undefined);
      const loader = makeLoader("a.wasm");
      await loader.ensureLoaded();

      loader.updateOptions({ timeoutMs: 5_000, maxRetries: 5 });
      expect(loader.isReady).toBe(true);
      expect(warn).not.toHaveBeenCalled();

      await loader.ensureLoaded();
      expect(initWasmMock).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });

  it("treats an equal URL object as unchanged (compares by string, not reference)", async () => {
    // URL 对象每次构造都是新引用：若 updateOptions 用 !== 比较，同一地址
    // 会被反复判为"变更"，已 ready 的 loader 被不断 reset 并刷幂等警告。
    // 必须按 String 归一化比较（与 export.worker.ts 的 loadedWasmKey 一致）。
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      initWasmMock.mockResolvedValue(undefined);
      const loader = makeLoader("https://cdn.example.com/x.wasm");
      await loader.ensureLoaded();
      expect(loader.isReady).toBe(true);

      loader.updateOptions({
        wasmUrl: new URL("https://cdn.example.com/x.wasm"),
      });
      expect(loader.isReady).toBe(true);
      expect(warn).not.toHaveBeenCalled();

      await loader.ensureLoaded();
      expect(initWasmMock).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });

  it("a superseded load must not stamp over the successor's ready state", async () => {
    // 时序：加载 A 的 attempt 1 失败进入 300ms 退避 → 退避期间换 URL
    // （reset）并完成新加载 B（ready）→ A 醒来跑 attempt 2。修复前
    // loadWithRetry 在每个 attempt 无条件写 state="loading"，会把 B 写好
    // 的 ready 覆盖掉且无人写回，isReady() 误报 false。
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      initWasmMock.mockRejectedValueOnce(new Error("old url fails"));
      const loader = makeLoader("a.wasm");
      const superseded = loader.ensureLoaded();
      // 等 attempt 1 确实发生（此后 A 处于退避等待中）。
      await vi.waitFor(() => expect(initWasmMock).toHaveBeenCalledTimes(1));

      loader.updateOptions({ wasmUrl: "b.wasm" });
      initWasmMock.mockResolvedValue(undefined);
      await loader.ensureLoaded();
      expect(loader.isReady).toBe(true);

      // A 的退避到期、attempt 2 跑完并 settle 之后，ready 必须保持。
      await superseded;
      expect(loader.isReady).toBe(true);
    } finally {
      warn.mockRestore();
    }
  });

  it("falls back to initWasm when the mocked modern-xlsx lacks initWasmSync", async () => {
    // The mock factory in this file has no initWasmSync export (test doubles
    // often won't): Node auto-init must step aside instead of throwing when
    // wasmUrl is unconfigured, so the initWasm path runs as before — with the
    // default URL (the shipped binary next to the entry) in place of the old
    // undefined.
    initWasmMock.mockResolvedValue(undefined);
    const loader = makeLoader();

    await loader.ensureLoaded();
    expect(loader.isReady).toBe(true);
    expect(initWasmMock).toHaveBeenCalledTimes(1);
    expect(String(initWasmMock.mock.calls[0][0])).toMatch(/modern-xlsx\.wasm$/);
  });
});
