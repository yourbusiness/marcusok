import { afterEach, describe, expect, it, vi } from "vitest";
import { configureWasm } from "@marcusok/xlsx-core";
import { exportInWorker, terminateWorker } from "../worker-exporter";

const sheets = [
  {
    name: "S",
    columns: [{ prop: "x", label: "X" }],
    data: [{ x: 1 }],
  },
];

class HealthyWorker {
  static instances: HealthyWorker[] = [];
  onmessage: ((e: { data: unknown }) => void) | null = null;
  onerror: ((err: { message?: string }) => void) | null = null;
  posted: { id: number }[] = [];

  constructor() {
    HealthyWorker.instances.push(this);
  }

  postMessage(data: { id: number }): void {
    this.posted.push(data);
  }

  terminate(): void {}
}

describe("exportInWorker worker construction failures", () => {
  afterEach(() => {
    terminateWorker();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    HealthyWorker.instances = [];
  });

  it("stops retrying a synchronous constructor failure until terminateWorker resets it", async () => {
    let constructions = 0;
    class ThrowingWorker {
      constructor() {
        constructions++;
        throw new Error("SecurityError: cross-origin worker script");
      }
    }
    vi.stubGlobal("Worker", ThrowingWorker);
    vi.stubGlobal("window", {});
    configureWasm({ workerUrl: "https://cdn.example.com/export.worker.js" });
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});

    const options = { filename: "construction", download: false, sheets };
    const first = await exportInWorker(options, "workbook");
    const second = await exportInWorker(options, "workbook");

    expect(first.success).toBe(false);
    expect(second.success).toBe(false);
    // 固定 URL 的同步构造失败不应在每次导出时重复付出异常成本。
    expect(constructions).toBe(1);
    expect(warn).toHaveBeenCalledWith(expect.stringContaining("same-origin"));

    // 公开的回收逃生口必须解除永久禁用，否则换 URL 后仍无法恢复 worker。
    vi.stubGlobal("Worker", HealthyWorker);
    terminateWorker();
    const next = exportInWorker(options, "workbook");
    expect(HealthyWorker.instances).toHaveLength(1);
    const worker = HealthyWorker.instances[0];
    worker.onmessage!({
      data: {
        id: worker.posted[0].id,
        ok: true,
        bytes: new Uint8Array([1]),
        rowCount: 1,
      },
    });
    await expect(next).resolves.toMatchObject({ success: true });
  });
});
