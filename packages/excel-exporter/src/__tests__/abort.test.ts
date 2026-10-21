import { describe, expect, it } from "vitest";
import { exportExcel } from "../index";
import { exportAsStream } from "../streaming-builder";

const sheets = [
  {
    name: "S",
    columns: [{ prop: "a", label: "A" }],
    data: [{ a: 1 }],
  },
];

describe("AbortSignal cancellation", () => {
  it("pre-aborted main export resolves a structured AbortError with 0 -> 1 progress", async () => {
    const ac = new AbortController();
    ac.abort();
    const progress: number[] = [];
    const r = await exportExcel({
      filename: "abort-main",
      download: false,
      mode: "main",
      signal: ac.signal,
      onProgress: (p) => progress.push(p),
      sheets,
    });
    expect(r.success).toBe(false);
    expect(r.error?.name).toBe("AbortError");
    // Cancellation happens before any engine work; no degradation ran.
    expect(r.engine).toBeUndefined();
    expect(progress).toEqual([0, 1]);
  });

  it("pre-aborted stream export resolves the same failure shape", async () => {
    const ac = new AbortController();
    ac.abort();
    const r = await exportExcel({
      filename: "abort-stream",
      download: false,
      mode: "stream",
      signal: ac.signal,
      sheets,
    });
    expect(r.success).toBe(false);
    expect(r.error?.name).toBe("AbortError");
    expect(r.engine).toBeUndefined();
  });

  it("exportAsStream honors an already-aborted signal", async () => {
    const ac = new AbortController();
    ac.abort();
    await expect(
      exportAsStream(sheets, undefined, ac.signal),
    ).rejects.toMatchObject({ name: "AbortError" });
  });
});
