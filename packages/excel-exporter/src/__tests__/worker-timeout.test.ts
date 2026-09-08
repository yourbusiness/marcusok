import { afterEach, describe, expect, it, vi } from "vitest";
import { exportInWorker, terminateWorker } from "../worker-exporter";
import { configureWasm } from "../wasm-loader";

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
    columns: [{ key: "x", header: "X" }],
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
});
