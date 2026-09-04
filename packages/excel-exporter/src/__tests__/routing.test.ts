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

  it("worker failure retries on the main thread, preserving styles (terminal 1 exactly once)", async () => {
    stubBrowserWorkerEnv();
    // exportInWorker catches worker errors and resolves with success:false;
    // exportExcel now retries on the main thread (modern-xlsx, styles kept)
    // instead of degrading straight to SheetJS. WASM was sync-initialized by
    // ./setup, so the retry succeeds.
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
      expect(r.engine).toBe("modern-xlsx");
      expect(r.mode).toBe("main");
      expect(r.success).toBe(true);
      expect(progress).toEqual([0, 1]);
    } finally {
      vi.unstubAllGlobals();
      vi.mocked(exportInWorker).mockReset();
    }
  });

  it("worker failure on the stream route retries as a main-thread fast stream", async () => {
    stubBrowserWorkerEnv();
    // Explicit stream mode in a browser routes through the worker; the retry
    // runs fast-xlsx on the main thread, which needs no WASM at all.
    vi.mocked(exportInWorker).mockResolvedValue({
      success: false,
      error: new Error("worker boom"),
    });
    try {
      const r = await exportExcel({
        filename: "worker-stream-retry",
        download: false,
        mode: "stream",
        sheets: [
          {
            name: "S",
            columns: [{ key: "x", header: "X" }],
            data: [{ x: 1 }],
          },
        ],
      });
      expect(r.success).toBe(true);
      expect(r.engine).toBe("modern-xlsx");
      expect(r.mode).toBe("stream");
    } finally {
      vi.unstubAllGlobals();
      vi.mocked(exportInWorker).mockReset();
    }
  });

  it("degrades to SheetJS only when the worker AND the main-thread retry both fail", async () => {
    stubBrowserWorkerEnv();
    vi.mocked(exportInWorker).mockResolvedValue({
      success: false,
      error: new Error("worker boom"),
    });
    // A format function that throws on every call, numbered per invocation:
    // the main-thread retry must throw "boom-1" and the SheetJS fallback
    // "boom-2", proving the full worker -> main -> SheetJS chain ran.
    let calls = 0;
    const sheets = [
      {
        name: "S",
        columns: [
          {
            key: "x",
            header: "X",
            format: () => {
              calls++;
              throw new Error(`boom-${calls}`);
            },
          },
        ],
        data: [{ x: 1 }],
      },
    ];
    try {
      const progress: number[] = [];
      const r = await exportExcel({
        filename: "worker-chain-fail",
        download: false,
        mode: "worker",
        sheets,
        onProgress: (p) => progress.push(p),
      });
      expect(r.success).toBe(false);
      expect(r.error?.message).toBe("boom-2");
      expect(progress).toEqual([0, 1]);
    } finally {
      vi.unstubAllGlobals();
      vi.mocked(exportInWorker).mockReset();
    }
  });
});
