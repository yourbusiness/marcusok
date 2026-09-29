import { describe, expect, it } from "vitest";

// No mocks: exercises the real zero-configuration Node path through core's
// public surface — the loader must locate the shipped wasm on disk
// (import.meta.url-relative, symlink-safe) and initialize it synchronously,
// and the engine re-export surface must actually work end to end.
import { getWasmLoader, Workbook, readBuffer } from "../index";

describe("core engine surface (integration, real wasm binary)", () => {
  it("loads the real modern-xlsx.wasm with zero configuration", async () => {
    const loader = getWasmLoader();
    await loader.ensureLoaded();
    expect(loader.isReady).toBe(true);
  });

  it("reads a workbook through the re-exported readBuffer", async () => {
    const wb = new Workbook();
    wb.addSheet("Core").cell("A1").value = "hello";
    const bytes = await wb.toBuffer();

    const back = await readBuffer(new Uint8Array(bytes));
    expect(back.getSheet("Core")?.cell("A1").value).toBe("hello");
  });
});
