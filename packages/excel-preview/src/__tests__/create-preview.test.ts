// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest";
import { runInNewContext } from "node:vm";
import { createPreview } from "../index";

/**
 * createPreview 初始 sheet 选项的校验语义：
 *  - 无效（名称不存在 / 索引越界或非整数）经 onError 上报（与实例方法
 *    setSheet 同 message、同 code），随后仍回退渲染文件的 activeSheet——
 *    修复前是纯静默回退，拼错的 sheet 名无声丢失；
 *  - 有效值不触发 onError。
 *
 * 观测手段沿用 abort.test.ts 的 CSV 技巧：happy-dom 无 Worker 全局 →
 * 主线程解析，CSV 字节产出单 sheet（名 "CSV"），boot 的异步链在几十毫秒内
 * 走完。
 */
const csvBytes = new TextEncoder().encode("a,b\n1,2\n");

async function flushBoot(): Promise<void> {
  // 给 boot 的微任务链（toBytes → 主线程解析 → 同步渲染）走完的余量
  await new Promise((res) => setTimeout(res, 20));
}

describe("createPreview 初始 sheet 校验", () => {
  it("无效名称：onError 报 Sheet not found，仍回退渲染 activeSheet", async () => {
    const onParsed = vi.fn();
    const onError = vi.fn();
    const container = document.createElement("div");
    document.body.appendChild(container);

    createPreview(container, {
      source: csvBytes,
      sheet: "NoSuchSheet",
      onParsed,
      onError,
    });
    await flushBoot();

    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError).toHaveBeenCalledWith({
      code: "UNKNOWN",
      message: "Sheet not found: NoSuchSheet",
    });
    // 回退渲染照常完成：onParsed 触发且渲染的是唯一 sheet
    expect(onParsed).toHaveBeenCalledTimes(1);
    expect(onParsed.mock.calls[0][0].sheetNames).toEqual(["CSV"]);
    expect(container.querySelectorAll(".xpv-cell").length).toBeGreaterThan(0);
  });

  it("越界/非整数索引：同口径 onError 后回退", async () => {
    for (const bad of [99, -1, 1.5]) {
      const onParsed = vi.fn();
      const onError = vi.fn();
      const container = document.createElement("div");
      document.body.appendChild(container);
      createPreview(container, {
        source: csvBytes,
        sheet: bad,
        onParsed,
        onError,
      });
      await flushBoot();
      expect(onError).toHaveBeenCalledTimes(1);
      expect(onError).toHaveBeenCalledWith({
        code: "UNKNOWN",
        message: `Sheet not found: ${String(bad)}`,
      });
      expect(onParsed).toHaveBeenCalledTimes(1);
    }
  });

  it("接受跨 realm 的 Uint8Array 源（不依赖 instanceof）", async () => {
    const n = [...csvBytes];
    const foreign = runInNewContext("new Uint8Array(n)", { n }) as Uint8Array;
    expect(foreign instanceof Uint8Array).toBe(false);
    const onParsed = vi.fn();
    const onError = vi.fn();
    const container = document.createElement("div");
    document.body.appendChild(container);
    createPreview(container, {
      source: foreign,
      onParsed,
      onError,
    });
    await flushBoot();
    expect(onError).not.toHaveBeenCalled();
    expect(onParsed).toHaveBeenCalledTimes(1);
  });

  it("异步 onParsed 的 rejection 被兜住（不产生 unhandled rejection）", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const container = document.createElement("div");
      document.body.appendChild(container);
      createPreview(container, {
        source: csvBytes,
        // eslint-disable-next-line @typescript-eslint/no-misused-promises -- 有意模拟 async void 回调的 rejected promise
        onParsed: async () => {
          await Promise.resolve();
          throw new Error("async UI unmounted");
        },
      });
      await flushBoot();
      await new Promise((res) => setTimeout(res, 0));
      expect(
        warn.mock.calls.some((c) => String(c[0]).includes("onParsed")),
      ).toBe(true);
    } finally {
      warn.mockRestore();
    }
  });

  it("有效名称与有效索引：不触发 onError", async () => {
    for (const good of ["CSV", 0] as const) {
      const onParsed = vi.fn();
      const onError = vi.fn();
      const container = document.createElement("div");
      document.body.appendChild(container);
      createPreview(container, {
        source: csvBytes,
        sheet: good,
        onParsed,
        onError,
      });
      await flushBoot();
      expect(onError).not.toHaveBeenCalled();
      expect(onParsed).toHaveBeenCalledTimes(1);
      expect(onParsed.mock.calls[0][0].sheetNames).toEqual(["CSV"]);
    }
  });

  it("实例方法 setSheet 的非整数/NaN 索引经 onError 上报，不抛裸异常", async () => {
    const onParsed = vi.fn();
    const onError = vi.fn();
    const container = document.createElement("div");
    document.body.appendChild(container);
    const preview = createPreview(container, {
      source: csvBytes,
      onParsed,
      onError,
    });
    await flushBoot();
    onError.mockClear();

    for (const bad of [1.5, NaN]) {
      expect(() => preview.setSheet(bad)).not.toThrow();
      expect(onError).toHaveBeenCalledWith({
        code: "UNKNOWN",
        message: `Sheet not found: ${String(bad)}`,
      });
    }
  });

  it("onParsed 抛错被隔离：不误报为解析错误，渲染继续", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const onError = vi.fn();
      const container = document.createElement("div");
      document.body.appendChild(container);
      createPreview(container, {
        source: csvBytes,
        onParsed: () => {
          throw new Error("user onParsed boom");
        },
        onError,
      });
      await flushBoot();
      expect(onError).not.toHaveBeenCalled();
      expect(warn).toHaveBeenCalledWith(
        expect.stringContaining("onParsed"),
        expect.any(Error),
      );
      expect(container.querySelectorAll(".xpv-cell").length).toBeGreaterThan(0);
    } finally {
      warn.mockRestore();
    }
  });

  it("onError 自身抛错不会产生未处理拒绝", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const container = document.createElement("div");
      document.body.appendChild(container);
      expect(() =>
        createPreview(container, {
          source: csvBytes,
          sheet: "NoSuchSheet",
          onError: () => {
            throw new Error("user onError boom");
          },
        }),
      ).not.toThrow();
      await flushBoot();
      expect(warn).toHaveBeenCalledWith(
        expect.stringContaining("onError"),
        expect.any(Error),
      );
    } finally {
      warn.mockRestore();
    }
  });

  it("未提供 sheet：不触发 onError（activeTab 默认）", async () => {
    const onParsed = vi.fn();
    const onError = vi.fn();
    const container = document.createElement("div");
    document.body.appendChild(container);
    createPreview(container, { source: csvBytes, onParsed, onError });
    await flushBoot();
    expect(onError).not.toHaveBeenCalled();
    expect(onParsed).toHaveBeenCalledTimes(1);
  });
});
