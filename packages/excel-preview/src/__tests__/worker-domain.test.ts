// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest";
import { strToU8 } from "fflate";
import { parseWorkbookSource } from "../worker/worker-client";

/**
 * worker 响应分流回归（P4）：domain 标记的域错误必须直传调用方，不回退
 * 主线程复跑；未标记的环境性失败仍回退。观测手段：喂给 worker 的字节是
 * CSV（主线程本可解析成功）——若 domain 错误仍触发回退，结果会变成 ok:true，
 * 与"直传错误"的期望立刻区分开。
 */
let preset: { ok: false; code: string; message: string; domain?: boolean };

class FakeWorker {
  onmessage: ((e: { data: unknown }) => void) | null = null;
  postMessage(msg: { id: number }) {
    const resp = preset;
    queueMicrotask(() => {
      this.onmessage?.({ data: { id: msg.id, ...resp } });
    });
  }
  terminate() {}
}
vi.stubGlobal("Worker", FakeWorker);

const csvBytes = strToU8("a,b\n1,2\n");

describe("worker 响应分流（domain 直传 / 非 domain 回退主线程）", () => {
  beforeEach(() => {
    preset = { ok: false, code: "UNKNOWN", message: "boom" };
  });

  it("domain 域错误直传：同样的字节不再回退主线程复跑", async () => {
    preset = {
      ok: false,
      code: "PASSWORD_PROTECTED",
      message: "protected",
      domain: true,
    };
    const r = await parseWorkbookSource({ bytes: csvBytes });
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.code).toBe("PASSWORD_PROTECTED");
      expect(r.message).toBe("protected");
    }
  });

  it("非 domain 失败回退主线程：CSV 在主线程解析成功", async () => {
    preset = { ok: false, code: "UNKNOWN", message: "worker exploded" };
    const r = await parseWorkbookSource({ bytes: csvBytes });
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.workbook.sheets[0].rowCount).toBe(2);
    }
  });
});
