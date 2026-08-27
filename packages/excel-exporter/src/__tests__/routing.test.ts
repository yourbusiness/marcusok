import { describe, it, expect, vi } from "vitest";
import { exportExcel } from "../index";
import { exportInWorker } from "../worker-exporter";
import { StylePresets } from "../style-presets";
// Node's fetch rejects file://, so sync-init the WASM (see setup.ts).
import "./setup";

// The browser worker route cannot run under vitest's node environment; mock
// the exporter boundary so the onProgress contract of the worker branch in
// exportExcel is still exercised (the Node-path cases below never reach it).
vi.mock("../worker-exporter", () => ({
  exportInWorker: vi.fn(),
}));

function stubBrowserWorkerEnv(): void {
  // pickMode treats a global Worker + window as "browser" -> worker route.
  vi.stubGlobal("Worker", class FakeWorker {});
  vi.stubGlobal("window", {});
}

// Node (vitest environment: 'node') has no Web Worker / window globals, so this
// suite verifies the env-aware fallbacks in pickMode without mocking.
describe("exportExcel mode routing (Node environment)", () => {
  it("forced 'worker' mode in a no-Worker env falls back to modern-xlsx, not SheetJS", async () => {
    const r = await exportExcel({
      filename: "routing-worker",
      download: false,
      mode: "worker",
      sheets: [
        {
          name: "S",
          columns: [{ key: "x", header: "X", style: StylePresets.currency }],
          data: [{ x: 1 }, { x: 2 }],
        },
      ],
    });
    // Previously: tried new Worker() -> threw -> degraded to SheetJS (styles
    // stripped). Now: main-thread Workbook, styles preserved.
    expect(r.success).toBe(true);
    expect(r.engine).toBe("modern-xlsx");
  });

  it("auto mode with a large dataset uses stream on the main thread in Node", async () => {
    const data = Array.from({ length: 60_000 }, (_, i) => ({ id: i }));
    const r = await exportExcel({
      filename: "routing-auto-stream",
      download: false,
      sheets: [{ name: "S", columns: [{ key: "id", header: "ID" }], data }],
    });
    expect(r.success).toBe(true);
    expect(r.engine).toBe("modern-xlsx");
    expect(r.mode).toBe("stream");
  });

  it("emits the documented 0 -> 1 onProgress pair even on the SheetJS fallback", async () => {
    // Force the early-bail fallback (WASM reported unsupported). The fallback
    // itself never reports progress; exportExcel must still open and close
    // the sequence exactly once each (types.ts onProgress contract).
    vi.stubGlobal("WebAssembly", undefined);
    try {
      const progress: number[] = [];
      const r = await exportExcel({
        filename: "routing-fallback-progress",
        download: false,
        sheets: [
          {
            name: "S",
            columns: [{ key: "x", header: "X" }],
            data: [{ x: 1 }, { x: 2 }],
          },
        ],
        onProgress: (p) => progress.push(p),
      });
      expect(r.engine).toBe("sheetjs");
      expect(r.success).toBe(true);
      expect(progress).toEqual([0, 1]);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

describe("exportExcel worker branch onProgress contract (mocked worker)", () => {
  it("success route emits the terminal 1 exactly once", async () => {
    stubBrowserWorkerEnv();
    vi.mocked(exportInWorker).mockResolvedValue({
      success: true,
      blob: new Blob(["x"]),
      engine: "modern-xlsx",
      mode: "worker",
    });
    try {
      const progress: number[] = [];
      const r = await exportExcel({
        filename: "worker-progress-ok",
        download: false,
        // Explicit worker mode: auto would route a 1-row browser export to
        // the main thread (< WORKER_THRESHOLD) and never reach the worker
        // branch under test.
        mode: "worker",
        sheets: [
          {
            name: "S",
            columns: [{ key: "x", header: "X" }],
            data: [{ x: 1 }],
          },
        ],
        onProgress: (p) => progress.push(p),
      });
      expect(r.success).toBe(true);
      expect(exportInWorker).toHaveBeenCalledWith(
        expect.objectContaining({ filename: "worker-progress-ok" }),
        "workbook",
      );
      expect(progress).toEqual([0, 1]);
    } finally {
      vi.unstubAllGlobals();
      vi.mocked(exportInWorker).mockReset();
    }
  });

  it("failure -> SheetJS fallback route emits the terminal 1 exactly once (was [0, 1, 1])", async () => {
    stubBrowserWorkerEnv();
    // exportInWorker catches worker errors and resolves with success:false;
    // exportExcel then degrades to SheetJS. Pre-fix, the terminal 1 was
    // emitted before checking success AND again in finishWithSheetJS's
    // finally, violating the types.ts "exactly once" contract.
    vi.mocked(exportInWorker).mockResolvedValue({
      success: false,
      error: new Error("worker boom"),
    });
    try {
      const progress: number[] = [];
      const r = await exportExcel({
        filename: "worker-progress-fail",
        download: false,
        mode: "worker", // same reason as above: reach the worker branch
        sheets: [
          {
            name: "S",
            columns: [{ key: "x", header: "X" }],
            data: [{ x: 1 }],
          },
        ],
        onProgress: (p) => progress.push(p),
      });
      expect(r.engine).toBe("sheetjs");
      expect(r.success).toBe(true);
      expect(progress).toEqual([0, 1]);
    } finally {
      vi.unstubAllGlobals();
      vi.mocked(exportInWorker).mockReset();
    }
  });
});
