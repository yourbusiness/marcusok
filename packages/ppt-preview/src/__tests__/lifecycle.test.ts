// @vitest-environment happy-dom
/**
 * 生命周期测试：加载中 destroy、迟到 resolve/reject 的作废语义、失败终态、
 * 重复 destroy、destroy 后回调静默。引擎（@silurus/ooxml/pptx）全 mock，
 * 时序用受控 deferred 推进。
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const loadMock = vi.hoisted(() => vi.fn());
const renderSlideMock = vi.hoisted(() => vi.fn());
const engineDestroyMock = vi.hoisted(() => vi.fn());
vi.mock("@silurus/ooxml/pptx", () => ({
  PptxPresentation: {
    load: (...args: unknown[]) => loadMock(...args),
  },
  buildPptxTextLayer: vi.fn(),
}));

import { createPptPreview } from "../render/render";
import { simplePptx } from "./fixtures";

/** 受控 deferred：让 load 的 resolve/reject 时机由测试驱动。 */
function deferred<T>(): {
  promise: Promise<T>;
  resolve: (v: T) => void;
  reject: (e: unknown) => void;
} {
  let resolve!: (v: T) => void;
  let reject!: (e: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

function fakeEngine(slides = 2) {
  return {
    slideCount: slides,
    slideWidth: 12192000,
    slideHeight: 6858000,
    renderSlide: (...args: unknown[]) => renderSlideMock(...args),
    destroy: engineDestroyMock,
  };
}

function magicSource(): Uint8Array {
  return simplePptx();
}

let container: HTMLElement;

beforeEach(() => {
  document.body.textContent = "";
  container = document.createElement("div");
  document.body.appendChild(container);
  document.getElementById("marcusok-ppt-preview-styles")?.remove();
  loadMock.mockReset();
  renderSlideMock.mockReset();
  engineDestroyMock.mockReset();
  renderSlideMock.mockResolvedValue(undefined);
});

describe("加载生命周期", () => {
  it("正常加载：占位建立、slideCount 可读", async () => {
    loadMock.mockResolvedValue(fakeEngine(3));
    const inst = createPptPreview(container, { source: magicSource() });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    expect(inst.slideCount).toBe(3);
    inst.destroy();
  });

  it("加载中 destroy：旧结果作废，迟到 resolve 的引擎立即被销毁、不上报错误", async () => {
    const d = deferred<ReturnType<typeof fakeEngine>>();
    loadMock.mockReturnValue(d.promise);
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: magicSource(),
      onError,
    });
    // 等 load 已发出（pendingLoad 非 null）再 destroy——load 尚未发出时的
    // destroy 走 boot 检查点提前退出，不覆盖"作废在途 load"路径
    await vi.waitFor(() => {
      expect(loadMock).toHaveBeenCalledTimes(1);
    });
    inst.destroy();
    expect(engineDestroyMock).not.toHaveBeenCalled(); // 引擎尚未到手
    d.resolve(fakeEngine(2)); // 迟到 resolve
    await vi.waitFor(() => {
      expect(engineDestroyMock).toHaveBeenCalledTimes(1);
    });
    expect(onError).not.toHaveBeenCalled(); // 取消不是错误
    expect(container.children).toHaveLength(0);
  });

  it("加载中 destroy 后迟到 reject：静默（不触发 onError）", async () => {
    const d = deferred<ReturnType<typeof fakeEngine>>();
    loadMock.mockReturnValue(d.promise);
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: magicSource(),
      onError,
    });
    await vi.waitFor(() => {
      expect(loadMock).toHaveBeenCalledTimes(1);
    });
    inst.destroy();
    d.reject(new Error("aborted"));
    await new Promise((r) => setTimeout(r, 0));
    expect(onError).not.toHaveBeenCalled();
    expect(engineDestroyMock).not.toHaveBeenCalled();
  });

  it("load 同步抛错：归一上报，不产生 unhandled rejection", async () => {
    // 上游 load() 正常返回 Promise，但非法选项等编程期错误可能同步抛出；
    // 该调用在 try 之外时 void boot() 会变成 unhandled rejection，用户无反馈。
    loadMock.mockImplementation(() => {
      throw Object.assign(new Error("unsupported option"), {
        code: "unsupported-content",
      });
    });
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: magicSource(),
      onError,
    });
    await vi.waitFor(() => {
      expect(onError).toHaveBeenCalledTimes(1);
    });
    expect(onError.mock.calls[0][0].code).toBe("UNSUPPORTED");
    expect(container.children).toHaveLength(0);
    inst.destroy();
  });

  it("load 发出前 destroy：boot 在检查点退出，load 永不发出", async () => {
    loadMock.mockResolvedValue(fakeEngine());
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: magicSource(),
      onError,
    });
    inst.destroy(); // 同步立即销毁（boot 还在 toBytes 的 await 处）
    await new Promise((r) => setTimeout(r, 10));
    expect(loadMock).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
  });

  it("加载 reject（未 destroy）：onError 上报归一错误、容器清空", async () => {
    loadMock.mockRejectedValue(
      Object.assign(new Error("The supplied password is incorrect."), {
        code: "invalid-password",
      }),
    );
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: magicSource(),
      onError,
    });
    await vi.waitFor(() => {
      expect(onError).toHaveBeenCalledTimes(1);
    });
    expect(onError.mock.calls[0][0].code).toBe("PASSWORD_PROTECTED");
    expect(container.children).toHaveLength(0);
    expect(inst.slideCount).toBe(0); // 失败终态作废页数
    inst.destroy(); // 失败后的 destroy 幂等安全
  });

  it("sniff 拦截（非 ZIP 非 OLE2）不触发引擎 load", async () => {
    loadMock.mockResolvedValue(fakeEngine());
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]),
      onError,
    });
    await vi.waitFor(() => {
      expect(onError).toHaveBeenCalledTimes(1);
    });
    expect(onError.mock.calls[0][0].code).toBe("UNSUPPORTED");
    expect(loadMock).not.toHaveBeenCalled();
    inst.destroy();
  });

  it("空 deck：slideCount=0、不触发 onRendered/onError", async () => {
    loadMock.mockResolvedValue(fakeEngine(0));
    const onRendered = vi.fn();
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: magicSource(),
      onRendered,
      onError,
    });
    await vi.waitFor(() => {
      expect(loadMock).toHaveBeenCalledTimes(1);
    });
    await new Promise((r) => setTimeout(r, 10));
    expect(inst.slideCount).toBe(0);
    expect(onRendered).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
    inst.destroy();
  });

  it("deck 尺寸异常（0 宽）→ CORRUPT，不建占位", async () => {
    const bad = fakeEngine(2);
    Object.assign(bad, { slideWidth: 0, slideHeight: 6858000 });
    loadMock.mockResolvedValue(bad);
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: magicSource(),
      onError,
    });
    await vi.waitFor(() => {
      expect(onError).toHaveBeenCalledTimes(1);
    });
    expect(onError.mock.calls[0][0].code).toBe("CORRUPT");
    expect(container.children).toHaveLength(0);
    expect(inst.slideCount).toBe(0);
  });

  it("重复 destroy 幂等；destroy 后回调静默", async () => {
    loadMock.mockResolvedValue(fakeEngine(2));
    const onRendered = vi.fn();
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: magicSource(),
      onRendered,
      onError,
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(2);
    });
    inst.destroy();
    inst.destroy(); // 幂等
    expect(engineDestroyMock).toHaveBeenCalledTimes(1);
    expect(container.children).toHaveLength(0);
    // destroy 后渲染失败不再上报（在途渲染的迟到 reject 由 renderer 静默）
    expect(onError).not.toHaveBeenCalled();
    expect(onRendered).not.toHaveBeenCalled();
  });

  it("boot 后半段同步初始化抛异常（IO 缺失环境）→ UNKNOWN 上报 + 清场，不产生 unhandled rejection", async () => {
    // 模拟无 IntersectionObserver 的环境（极旧 WebView / SSR 误用）：
    // load 成功后 LazySlideRenderer 构造会同步抛——boot 后半段必须有
    // 兜底（否则 void boot() 变 unhandled rejection，用户无任何反馈）
    vi.stubGlobal("IntersectionObserver", undefined);
    try {
      loadMock.mockResolvedValue(fakeEngine(2));
      const onError = vi.fn();
      createPptPreview(container, { source: magicSource(), onError });
      await vi.waitFor(() => {
        expect(onError).toHaveBeenCalledTimes(1);
      });
      expect(onError.mock.calls[0][0].code).toBe("UNKNOWN");
      // fail 清场：引擎已到手则销毁、容器清空（失败终态可再 destroy）
      expect(engineDestroyMock).toHaveBeenCalledTimes(1);
      expect(container.children).toHaveLength(0);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("参数校验同步抛 TypeError（不进异步回调）", () => {
    loadMock.mockResolvedValue(fakeEngine());
    expect(() =>
      createPptPreview(container, { source: magicSource(), slideWidth: 0 }),
    ).toThrow(TypeError);
    expect(() =>
      createPptPreview(container, {
        source: magicSource(),
        slideWidth: Number.NaN,
      }),
    ).toThrow(TypeError);
    expect(() =>
      createPptPreview(container, { source: magicSource(), dpr: -1 }),
    ).toThrow(TypeError);
    expect(loadMock).not.toHaveBeenCalled();
  });
});
