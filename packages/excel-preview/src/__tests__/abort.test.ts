// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest";
import { parseWorkbookSource } from "../worker/worker-client";
import { parseWorkbookBytes } from "../index";

/**
 * 解析取消传播（ABORTED）回归：
 *  - 取消结果必须直传调用方，不回退主线程复跑（复跑违背取消意图）；
 *  - createPreview 的 destroy 静默（取消不是错误，不触发 onError）。
 *
 * 观测手段沿用 worker-domain.test.ts 的 CSV 技巧：喂给解析的字节是 CSV，
 * 主线程本可解析成功——若 ABORTED 误走主线程回退，结果会变成 ok:true，
 * 与"取消直传"的期望立刻区分开。
 */
const csvBytes = new TextEncoder().encode("a,b\n1,2\n");

describe("parseWorkbookSource 取消（主线程路径）", () => {
  it("已中止的 signal 直接短路：ABORTED，且不进主线程解析（CSV 未被解析）", async () => {
    // happy-dom 无 Worker 全局 → getOrCreateWorker 返回 null → 主线程路径；
    // 入口短路先于路径选择，sniff/readBuffer 均不执行
    const r = await parseWorkbookSource({
      bytes: csvBytes,
      signal: AbortSignal.abort(),
    });
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.code).toBe("ABORTED");
      expect(r.message).toContain("cancelled");
    }
  });

  it("解析成功路径不受 signal 存在的影响（未中止即正常解析）", async () => {
    const controller = new AbortController();
    const r = await parseWorkbookSource({
      bytes: csvBytes,
      signal: controller.signal,
    });
    expect(r.ok).toBe(true);
  });
});

describe("parseWorkbookSource 取消（worker 等待中 abort）", () => {
  it("等待 worker 响应期间 abort：立即返回 ABORTED，不回退主线程", async () => {
    vi.resetModules();
    // 挂死 worker：postMessage 吞掉消息永不响应（abort 生效前请求悬置）
    class HungWorker {
      onmessage: ((e: { data: unknown }) => void) | null = null;
      postMessage() {}
      terminate() {}
    }
    vi.stubGlobal("Worker", HungWorker);
    const { parseWorkbookSource: fresh } =
      await import("../worker/worker-client");

    const controller = new AbortController();
    const promise = fresh({ bytes: csvBytes, signal: controller.signal });
    // 同步 abort：请求已发出（postMessage 已过）、响应未到
    controller.abort();
    const r = await promise;
    expect(r.ok).toBe(false);
    if (!r.ok) {
      // 若 ABORTED 误走回退链，CSV 会以 ok:true 解析成功（CSV 技巧反证）
      expect(r.code).toBe("ABORTED");
    }
  });

  it("abort 摘除监听器：settle 后再 abort 不产生悬挂引用/副作用", async () => {
    vi.resetModules();
    class EchoWorker {
      onmessage: ((e: { data: unknown }) => void) | null = null;
      postMessage(msg: { id: number }) {
        queueMicrotask(() => {
          this.onmessage?.({
            // CSV 字节交 worker 假装"解析失败"（非 domain），驱动主线程回退
            // 成功——覆盖 settle 分支后 removeEventListener 的执行路径
            data: { id: msg.id, ok: false, code: "UNKNOWN", message: "x" },
          });
        });
      }
      terminate() {}
    }
    vi.stubGlobal("Worker", EchoWorker);
    const { parseWorkbookSource: fresh } =
      await import("../worker/worker-client");

    const controller = new AbortController();
    // 正常完成（worker 失败 → 主线程回退解析 CSV 成功）
    const r = await fresh({ bytes: csvBytes, signal: controller.signal });
    expect(r.ok).toBe(true);
    // settle 后 abort：Promise 已返回，无任何可观测副作用（不抛、不挂）
    controller.abort();
    await new Promise((res) => setTimeout(res, 0));
  });
});

describe("parseWorkbookBytes 取消（公共低层 API）", () => {
  it("已中止 signal：抛带 .code === ABORTED 的 Error", async () => {
    await expect(
      parseWorkbookBytes(csvBytes, { signal: AbortSignal.abort() }),
    ).rejects.toMatchObject({ code: "ABORTED" });
  });
});

describe("createPreview destroy 取消", () => {
  it("解析等待期间 destroy：取消静默——onError/onParsed 均不触发", async () => {
    vi.resetModules();
    // 挂死 worker 让 boot 悬在等待：没有 abort 机制时它会挂满 120s 超时
    class HungWorker {
      onmessage: ((e: { data: unknown }) => void) | null = null;
      postMessage() {}
      terminate() {}
    }
    vi.stubGlobal("Worker", HungWorker);
    const { createPreview: freshCreate } = await import("../index");

    const onParsed = vi.fn();
    const onError = vi.fn();
    const container = document.createElement("div");
    document.body.appendChild(container);
    const preview = freshCreate(container, {
      source: csvBytes,
      onParsed,
      onError,
    });
    preview.destroy();
    // 给 boot 的微任务链走完的余量：abort 应让 boot 立刻静默返回
    await new Promise((res) => setTimeout(res, 20));
    expect(onParsed).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
    expect(container.textContent).toBe("");
  });
});
