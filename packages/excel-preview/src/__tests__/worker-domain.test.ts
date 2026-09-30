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

describe("worker 超时（挂死保护）", () => {
  // 挂死 worker：postMessage 后永不响应，模拟 wasm 初始化悬置/引擎死循环。
  // 观测手段沿用本文件的 CSV 技巧：字节是 CSV，主线程本可解析成功——若
  // 超时错误误走主线程回退，结果会变成 ok:true，与"不回退"的期望立刻区分。
  let terminations = 0;

  class HungWorker {
    onmessage: ((e: { data: unknown }) => void) | null = null;
    postMessage() {}
    terminate() {
      terminations++;
    }
  }

  it("超时终止共享 worker、拒绝在途请求且不回退主线程", async () => {
    // 模块级 worker 单例可能已被上方用例占用：重置模块取全新副本，
    // 保证 getOrCreateWorker 创建的是本次的 HungWorker
    vi.resetModules();
    vi.useFakeTimers();
    vi.stubGlobal("Worker", HungWorker);
    const { parseWorkbookSource: fresh } =
      await import("../worker/worker-client");

    const promise = fresh({ bytes: csvBytes });
    await vi.advanceTimersByTimeAsync(120_000);
    const r = await promise;
    expect(terminations).toBe(1);
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.noFallback).toBe(true);
      expect(r.message).toContain("timed out");
    }
    vi.useRealTimers();
  });

  it("超时后不永久禁用 worker 路径：下次请求重建新实例", async () => {
    vi.useFakeTimers();
    const { parseWorkbookSource: fresh } =
      await import("../worker/worker-client");

    const promise = fresh({ bytes: csvBytes });
    await vi.advanceTimersByTimeAsync(120_000);
    await promise;
    // 第二轮超时再次 terminate 的是新实例（workerBroken 未被置位）
    expect(terminations).toBe(2);
    vi.useRealTimers();
  });
});

describe("worker 构造失败（跨域脚本同步抛错）", () => {
  // worker 脚本受浏览器同源限制：parseWorkerUrl 配成裸 CDN 域名时
  // new Worker() 在 Chrome 下同步抛 SecurityError。异常若穿透到调用方，
  // 主线程回退不会生效——必须就地吞掉、置 workerBroken 并回退。
  // 观测手段沿用 CSV 技巧：字节是 CSV，主线程本可解析成功。
  it("构造异常回退主线程（不 reject 调用方），且置 workerBroken 不再重试构造", async () => {
    vi.resetModules();
    let constructions = 0;
    class CrossOriginWorker {
      onmessage: ((e: { data: unknown }) => void) | null = null;
      constructor() {
        constructions++;
        throw new Error(
          "Failed to construct 'Worker': Script cannot be accessed from origin",
        );
      }
      postMessage() {}
      terminate() {}
    }
    vi.stubGlobal("Worker", CrossOriginWorker);
    const { parseWorkbookSource: fresh } =
      await import("../worker/worker-client");

    const r1 = await fresh({ bytes: csvBytes });
    expect(r1.ok).toBe(true); // 主线程回退生效（CSV 可解析）
    if (r1.ok) expect(r1.workbook.sheets[0].rowCount).toBe(2);
    const r2 = await fresh({ bytes: csvBytes });
    expect(r2.ok).toBe(true);
    // workerBroken 已置位：第二次请求直接走主线程，不再重付构造
    expect(constructions).toBe(1);
  });
});
