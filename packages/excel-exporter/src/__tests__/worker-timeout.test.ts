import { afterEach, describe, expect, it, vi } from "vitest";
import { exportInWorker, terminateWorker } from "../worker-exporter";
import { configureWasm } from "../wasm-loader";
import type { ExportOptions } from "../types";

// A fake Worker whose postMessage never produces a response, driving every
// request into the 120s timeout without waiting real time (vi.useFakeTimers).
class FakeWorker {
  static instances: FakeWorker[] = [];
  onmessage: ((e: unknown) => void) | null = null;
  onerror: ((err: { message?: string }) => void) | null = null;
  terminated = false;

  constructor() {
    FakeWorker.instances.push(this);
  }

  postMessage(..._args: unknown[]): void {}
  terminate(): void {
    this.terminated = true;
  }
}

function stubBrowser(): void {
  vi.stubGlobal("Worker", FakeWorker);
  vi.stubGlobal("window", {});
}

const sheets = [
  {
    name: "S",
    columns: [{ prop: "x", label: "X" }],
    data: [{ x: 1 }],
  },
];

describe("exportInWorker timeout recovery", () => {
  afterEach(() => {
    terminateWorker();
    vi.useRealTimers();
    vi.unstubAllGlobals();
    FakeWorker.instances = [];
  });

  it("terminates the timed-out worker and drops the cached instance", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();

    // FakeWorker.postMessage never responds -> the timeout fires for the only
    // pending request. Previously the wedged worker stayed cached forever and
    // every later export queued behind it.
    const first = exportInWorker(
      { filename: "timeout-1", download: false, sheets },
      "workbook",
    );
    await vi.advanceTimersByTimeAsync(120_000);
    const r1 = await first;

    expect(r1.success).toBe(false);
    expect(r1.error?.message).toMatch(/timed out after 120000ms/);
    expect(FakeWorker.instances[0].terminated).toBe(true);

    // The cached reference was dropped, so the next export gets a fresh worker
    // (which also times out here, proving it was actually created).
    const second = exportInWorker(
      { filename: "timeout-2", download: false, sheets },
      "workbook",
    );
    expect(FakeWorker.instances.length).toBe(2);
    await vi.advanceTimersByTimeAsync(120_000);
    const r2 = await second;
    expect(r2.success).toBe(false);
    expect(FakeWorker.instances[1].terminated).toBe(true);
  });

  it("rejects sibling requests dispatched to the terminated worker", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();

    const a = exportInWorker(
      { filename: "sibling-a", download: false, sheets },
      "workbook",
    );
    const b = exportInWorker(
      { filename: "sibling-b", download: false, sheets },
      "workbook",
    );
    await vi.advanceTimersByTimeAsync(120_000);
    const [ra, rb] = await Promise.all([a, b]);

    // The timed-out request reports the timeout; its sibling on the same
    // worker is rejected (and its timer cleared) so its caller degrades to
    // the main-thread retry immediately instead of hanging for its own 120s.
    expect(ra.success).toBe(false);
    expect(ra.error?.message).toMatch(/timed out after 120000ms/);
    expect(rb.success).toBe(false);
    expect(rb.error?.message).toMatch(
      /terminated after a concurrent export timed out/,
    );
    expect(FakeWorker.instances[0].terminated).toBe(true);
  });

  it("cleans up immediately when postMessage throws (no delayed worker kill)", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();

    // postMessage throws synchronously on un-cloneable payloads (e.g. Symbol
    // or function values in row data -> DataCloneError).
    const throwing = vi
      .spyOn(FakeWorker.prototype, "postMessage")
      .mockImplementation(() => {
        throw new Error("DataCloneError: could not be cloned");
      });

    const r = await exportInWorker(
      { filename: "clone-error", download: false, sheets },
      "workbook",
    );
    expect(r.success).toBe(false);
    expect(r.error?.message).toMatch(/could not be cloned/);

    // Pre-fix the request stayed in `pending` with its 120s timer live, so
    // the timer later fired and terminated the healthy shared worker. The
    // worker must survive the full timeout window, and the next export must
    // reuse it (no new instance).
    await vi.advanceTimersByTimeAsync(120_000);
    expect(FakeWorker.instances[0].terminated).toBe(false);

    throwing.mockRestore();
    const second = exportInWorker(
      { filename: "after-clone-error", download: false, sheets },
      "workbook",
    );
    expect(FakeWorker.instances.length).toBe(1);
    await vi.advanceTimersByTimeAsync(120_000);
    const r2 = await second;
    expect(r2.success).toBe(false);
    expect(r2.error?.message).toMatch(/timed out after 120000ms/);
  });

  it("honors a custom workerTimeoutMs", async () => {
    stubBrowser();
    configureWasm({
      workerUrl: "/fake/export.worker.js",
      workerTimeoutMs: 5_000,
    });
    vi.useFakeTimers();
    try {
      const first = exportInWorker(
        { filename: "custom-timeout", download: false, sheets },
        "workbook",
      );
      await vi.advanceTimersByTimeAsync(5_000);
      const r1 = await first;
      expect(r1.success).toBe(false);
      expect(r1.error?.message).toMatch(/timed out after 5000ms/);
      expect(FakeWorker.instances[0].terminated).toBe(true);
    } finally {
      // Loader options merge across tests (module-level default loader);
      // restore the default so later tests see the 120s timeout.
      configureWasm({ workerTimeoutMs: 120_000 });
    }
  });
});

describe("exportInWorker error handling & payload stripping", () => {
  afterEach(() => {
    terminateWorker();
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    FakeWorker.instances = [];
  });

  it("onerror terminates the worker, drops the cached instance, and rejects only its own pending requests", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();

    // 两个请求派发到同一个共享 worker；脚本加载失败（如 404/CSP）时
    // ErrorEvent.message 通常为空串，错误折叠为兜底文案。
    const a = exportInWorker(
      { filename: "onerror-a", download: false, sheets },
      "workbook",
    );
    const b = exportInWorker(
      { filename: "onerror-b", download: false, sheets },
      "workbook",
    );
    expect(FakeWorker.instances.length).toBe(1);
    const w = FakeWorker.instances[0];
    w.onerror!({ message: "" });

    const [ra, rb] = await Promise.all([a, b]);
    expect(ra.success).toBe(false);
    expect(ra.error?.message).toBe("worker error");
    expect(rb.success).toBe(false);
    expect(w.terminated).toBe(true);

    // 缓存引用已丢弃：下一个导出创建全新 worker（而非复用已死实例）
    const c = exportInWorker(
      { filename: "onerror-c", download: false, sheets },
      "workbook",
    );
    expect(FakeWorker.instances.length).toBe(2);
    await vi.advanceTimersByTimeAsync(120_000);
    const rc = await c;
    expect(rc.success).toBe(false);
    expect(rc.error?.message).toMatch(/timed out after 120000ms/);
  });

  it("strips function-form formats (incl. nested columns) with a warning before postMessage", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});

    const posted: { options: ExportOptions }[] = [];
    vi.spyOn(FakeWorker.prototype, "postMessage").mockImplementation(
      (msg: unknown) => {
        posted.push(msg as { options: ExportOptions });
      },
    );

    const onProgress = (): void => {};
    const promise = exportInWorker(
      {
        filename: "strip-fn",
        download: false,
        onProgress, // 回调不可克隆，须随剥离一并移除
        sheets: [
          {
            name: "S",
            columns: [
              { prop: "a", label: "A", format: () => "X" },
              { prop: "b", label: "B", format: { type: "date" } },
              {
                label: "G",
                children: [{ prop: "c", label: "C", format: () => "Y" }],
              },
            ],
            data: [],
          },
        ],
      },
      "workbook",
    );
    await vi.advanceTimersByTimeAsync(120_000); // FakeWorker 不回应，走完超时
    const r = await promise;
    expect(r.success).toBe(false); // 超时失败即正常：这里只关心发出的载荷

    // types.ts 契约：函数 format 在 worker 路由被剥离并 console.warn，
    // FormatSpec（结构化）format 与列树嵌套位置都要处理正确。
    const cols = posted[0].options.sheets[0].columns;
    expect(cols[0].format).toBeUndefined();
    expect(cols[1].format).toEqual({ type: "date" });
    expect(cols[2].children![0].format).toBeUndefined();
    expect(posted[0].options.onProgress).toBeUndefined();
    expect(warn).toHaveBeenCalledWith(
      expect.stringContaining("function format"),
    );
  });
});
