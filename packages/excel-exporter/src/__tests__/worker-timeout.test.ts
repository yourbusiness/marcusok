import { afterEach, describe, expect, it, vi } from "vitest";
import { exportInWorker, terminateWorker } from "../worker-exporter";
import { configureWasm } from "@marcusok/xlsx-core";
import type { ExportOptions } from "../types";

// A fake Worker whose postMessage never produces a response, driving every
// request into the 120s timeout without waiting real time (vi.useFakeTimers).
// postMessage 载荷同步记录：载荷剥离（回调/overlay）的断言在不动 worker
// 协议 mock 的前提下即可锁定。
class FakeWorker {
  static instances: FakeWorker[] = [];
  onmessage: ((e: unknown) => void) | null = null;
  onerror: ((err: { message?: string }) => void) | null = null;
  terminated = false;
  posted: unknown[] = [];

  constructor() {
    FakeWorker.instances.push(this);
  }

  postMessage(data: unknown, ..._rest: unknown[]): void {
    this.posted.push(data);
  }
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

  it("strips the overlay option (DOM values would fail structured clone)", async () => {
    // overlay 由 exportExcel 包装层在主线程消费（遮罩已接线），但它是
    // ProgressOverlayOptions——合法携带 container: HTMLElement 一类不可
    // 克隆值。留在载荷里会让 worker 路由的每次 postMessage 同步抛
    // DataCloneError，20k+ 行导出全部退回主线程阻塞重建（worker 阈值要
    // 避免的正是这个），用户定制的遮罩还冻结在屏上。
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();

    const doomed = exportInWorker(
      {
        filename: "overlay-strip",
        download: false,
        sheets,
        // 非纯数据值：真实浏览器里 structuredClone 对它必抛 DataCloneError
        overlay: { container: { nodeType: 1 } as unknown as HTMLElement },
      },
      "workbook",
    );
    await vi.advanceTimersByTimeAsync(120_000);
    const r = await doomed;
    // FakeWorker.postMessage 不抛——超时路径证明请求确实发出去了；载荷
    // 断言才是本用例的目的：overlay 字段不得出现在 postMessage 参数里。
    expect(r.success).toBe(false);
    const payload = FakeWorker.instances[0].posted[0] as {
      options: Record<string, unknown>;
    };
    expect(payload.options).not.toHaveProperty("overlay");
    expect(payload.options).not.toHaveProperty("onProgress");
    expect(payload.options).toHaveProperty("filename");
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
    // 超时错误带 noFallback 标记：路由层据此跳过主线程重试（见 index.ts）
    expect(r1.error).toMatchObject({ noFallback: true });
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

  it("queues sibling requests and gives them a fresh worker after a timeout", async () => {
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
    // b is queued on the main thread, not dispatched to the wedged worker.
    expect(FakeWorker.instances.length).toBe(1);

    await vi.advanceTimersByTimeAsync(120_000);
    const ra = await a;
    expect(ra.success).toBe(false);
    expect(ra.error?.message).toMatch(/timed out after 120000ms/);
    expect(ra.error).toMatchObject({ noFallback: true });
    expect(FakeWorker.instances[0].terminated).toBe(true);

    // a's timeout dispatches b to a fresh worker and only THEN starts b's
    // clock, so the sibling is not failed just for waiting in the queue.
    expect(FakeWorker.instances.length).toBe(2);
    await vi.advanceTimersByTimeAsync(120_000);
    const rb = await b;
    expect(rb.success).toBe(false);
    expect(rb.error?.message).toMatch(/timed out after 120000ms/);
    expect(rb.error?.message).not.toMatch(/concurrent export timed out/);
    expect(rb.error).toMatchObject({ noFallback: true });
    expect(FakeWorker.instances[1].terminated).toBe(true);
  });

  it("aborts a queued request without touching the active one", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();

    const active = exportInWorker(
      { filename: "abort-active", download: false, sheets },
      "workbook",
    );
    const ac = new AbortController();
    const queued = exportInWorker(
      { filename: "abort-queued", download: false, sheets, signal: ac.signal },
      "workbook",
    );
    expect(FakeWorker.instances.length).toBe(1);

    ac.abort();
    const rq = await queued;
    expect(rq.success).toBe(false);
    expect(rq.error?.name).toBe("AbortError");
    expect(FakeWorker.instances[0].terminated).toBe(false);

    const w = FakeWorker.instances[0];
    const payload = w.posted[0] as { id: number };
    w.onmessage!({
      data: {
        id: payload.id,
        ok: true,
        bytes: new Uint8Array([1]),
        rowCount: 1,
      },
    });
    const ra = await active;
    expect(ra.success).toBe(true);
  });

  it("aborts the active request, terminates its worker, and re-dispatches the queue", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();

    const ac = new AbortController();
    const a = exportInWorker(
      { filename: "abort-a", download: false, sheets, signal: ac.signal },
      "workbook",
    );
    const b = exportInWorker(
      { filename: "abort-b", download: false, sheets },
      "workbook",
    );
    ac.abort();
    const ra = await a;
    expect(ra.success).toBe(false);
    expect(ra.error?.name).toBe("AbortError");
    expect(FakeWorker.instances[0].terminated).toBe(true);

    expect(FakeWorker.instances.length).toBe(2);
    const w2 = FakeWorker.instances[1];
    const payload = w2.posted[0] as { id: number };
    w2.onmessage!({
      data: {
        id: payload.id,
        ok: true,
        bytes: new Uint8Array([1]),
        rowCount: 1,
      },
    });
    const rb = await b;
    expect(rb.success).toBe(true);
  });

  it("never constructs a worker for an already-aborted signal", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    const ac = new AbortController();
    ac.abort();
    const r = await exportInWorker(
      { filename: "abort-first", download: false, sheets, signal: ac.signal },
      "workbook",
    );
    expect(r.success).toBe(false);
    expect(r.error?.name).toBe("AbortError");
    expect(FakeWorker.instances.length).toBe(0);
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

  it("onerror terminates the worker, drops the cached instance, and re-dispatches queued requests", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();

    // a 派发到共享 worker，b 在主线程排队；脚本加载失败（如 404/CSP）时
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

    const ra = await a;
    expect(ra.success).toBe(false);
    expect(ra.error?.message).toBe("worker error");
    expect(w.terminated).toBe(true);

    // b 从未派发到已死 worker：它被重新派发到新实例并可正常完成。
    expect(FakeWorker.instances.length).toBe(2);
    const w2 = FakeWorker.instances[1];
    const payload = w2.posted[0] as { id: number };
    w2.onmessage!({
      data: {
        id: payload.id,
        ok: true,
        bytes: new Uint8Array([1]),
        rowCount: 1,
      },
    });
    const rb = await b;
    expect(rb.success).toBe(true);
    expect(rb.rowCount).toBe(1);

    // 缓存引用已丢弃/替换：下一个导出复用健康的 w2，而不是再新建。
    const c = exportInWorker(
      { filename: "onerror-c", download: false, sheets },
      "workbook",
    );
    expect(FakeWorker.instances.length).toBe(2);
    await vi.advanceTimersByTimeAsync(120_000);
    const rc = await c;
    expect(rc.success).toBe(false);
    expect(rc.error?.message).toMatch(/timed out after 120000ms/);
    expect(w2.terminated).toBe(true);
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

  it("terminateWorker rejects in-flight requests as noFallback (与超时连坐同口径)", async () => {
    stubBrowser();
    configureWasm({ workerUrl: "/fake/export.worker.js" });
    vi.useFakeTimers();

    // FakeWorker 不回应 → 请求在途；terminateWorker 主动终止后必须直接失败
    // （路由层见 noFallback 即不再主线程重跑），不能退化成主线程重建。
    const promise = exportInWorker(
      { filename: "terminate-inflight", download: false, sheets },
      "workbook",
    );
    terminateWorker();
    const r = await promise;
    expect(r.success).toBe(false);
    expect((r.error as { noFallback?: true } | undefined)?.noFallback).toBe(
      true,
    );
    expect(r.error?.message).toMatch(/terminateWorker/);
    // 缓存引用已清：下一次导出创建全新 worker
    const next = exportInWorker(
      { filename: "terminate-next", download: false, sheets },
      "workbook",
    );
    await vi.advanceTimersByTimeAsync(120_000); // FakeWorker 依旧不回应，清尾
    await next;
    expect(FakeWorker.instances.length).toBe(2);
  });
});
