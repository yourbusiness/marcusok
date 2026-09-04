import { initWasm } from "modern-xlsx";
import type { ExportOptions } from "../types";
import { WorkbookBuilder } from "../workbook-builder";
import { exportAsStream } from "../streaming-builder";

interface WorkerRequest {
  id: number;
  options: ExportOptions;
  wasmUrl?: string | URL;
  mode: "workbook" | "stream";
}
interface WorkerResponse {
  id: number;
  ok: boolean;
  bytes?: Uint8Array;
  rowCount?: number;
  engine?: "modern-xlsx";
  error?: string;
  progress?: number;
}

// Track WASM initialization by the string form of the URL. modern-xlsx's
// initWasm is idempotent (first successful init wins; a later wasmUrl change
// cannot take effect in an already-initialized worker), so this key exists to
// avoid re-calling initWasm and re-reporting the "init" phase on every export
// -- comparing by string matters because a URL-typed wasmUrl arrives via
// structured clone as a fresh object each message and would never be `!==`.
// (wasmReady is separate from the key: an undefined wasmUrl maps to key null
// both before and after init, so the key alone cannot tell "never initialized".)
let wasmReady = false;
let loadedWasmKey: string | null = null;
self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const { id, options, wasmUrl, mode } = e.data;
  try {
    const urlKey = wasmUrl == null ? null : String(wasmUrl);
    if (mode !== "stream" && (!wasmReady || loadedWasmKey !== urlKey)) {
      const initStart = performance.now();
      await initWasm(wasmUrl);
      wasmReady = true;
      loadedWasmKey = urlKey;
      (self as unknown as Worker).postMessage({
        id,
        phase: "init",
        duration: performance.now() - initStart,
      });
    }

    const buildStart = performance.now();
    let bytes: Uint8Array;
    let rowCount: number;

    if (mode === "stream") {
      // Forward per-row progress to the main thread (throttled inside exportAsStream).
      const r = await exportAsStream(options.sheets, (progress) => {
        (self as unknown as Worker).postMessage({ id, progress });
      });
      bytes = r.bytes;
      rowCount = r.rowCount;
    } else {
      const builder = await WorkbookBuilder.create();
      for (const s of options.sheets) builder.addSheet(s);
      bytes = await builder.toBuffer();
      rowCount = options.sheets.reduce((sum, s) => sum + s.data.length, 0);
    }

    (self as unknown as Worker).postMessage({
      id,
      phase: "build",
      duration: performance.now() - buildStart,
    });

    const resp: WorkerResponse = {
      id,
      ok: true,
      bytes,
      rowCount,
      engine: "modern-xlsx",
    };
    (self as unknown as Worker).postMessage(resp, [bytes.buffer]);
  } catch (err) {
    const resp: WorkerResponse = {
      id,
      ok: false,
      error: (err as Error).message,
    };
    (self as unknown as Worker).postMessage(resp);
  }
};
