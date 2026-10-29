// @vitest-environment happy-dom
/**
 * 渲染管线测试：懒渲染队列、过期渲染、释放语义（无单页 cancel 句柄的
 * 引擎差异）、renderSlide 参数合同（skipMediaControls/单趟 onTextRun）、
 * 首张成功页 onRendered 语义、0-based 导航（含隐藏页的原始页序）。
 *
 * 引擎全 mock（@silurus/ooxml/pptx）；IO 用 FakeIO 驱动交叉回调；
 * happy-dom 无布局引擎，getBoundingClientRect 返回全 0——导航相关的
 * 可见性/滚动计算通过 spy 注入可控 rect。
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const loadMock = vi.hoisted(() => vi.fn());
const engineRenderMock = vi.hoisted(() => vi.fn());
const engineDestroyMock = vi.hoisted(() => vi.fn());
const buildTextLayerMock = vi.hoisted(() => vi.fn());
vi.mock("@silurus/ooxml/pptx", () => ({
  PptxPresentation: { load: (...a: unknown[]) => loadMock(...a) },
  buildPptxTextLayer: (...a: unknown[]) => buildTextLayerMock(...a),
}));

import { createPptPreview } from "../render/render";
import { buildPptx } from "./fixtures";

class FakeIO {
  static instances: FakeIO[] = [];
  callback: IntersectionObserverCallback;
  options: IntersectionObserverInit | undefined;
  constructor(
    cb: IntersectionObserverCallback,
    options?: IntersectionObserverInit,
  ) {
    this.callback = cb;
    this.options = options;
    FakeIO.instances.push(this);
  }
  observe(_el: Element) {}
  unobserve(_el: Element) {}
  disconnect() {}
  fire(el: Element, isIntersecting: boolean) {
    this.callback(
      [{ target: el, isIntersecting } as IntersectionObserverEntry],
      this as unknown as IntersectionObserver,
    );
  }
}

function fakeEngine(slides: number) {
  return {
    slideCount: slides,
    slideWidth: 12192000,
    slideHeight: 6858000,
    renderSlide: (...a: unknown[]) => engineRenderMock(...a),
    destroy: engineDestroyMock,
  };
}

/** 受控 deferred：挂起渲染以观察队列状态。 */
function deferred(): {
  promise: Promise<void>;
  resolve: () => void;
  reject: (e: unknown) => void;
} {
  let resolve!: () => void;
  let reject!: (e: unknown) => void;
  const promise = new Promise<void>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

let container: HTMLElement;

beforeEach(() => {
  document.body.textContent = "";
  container = document.createElement("div");
  document.body.appendChild(container);
  document.getElementById("marcusok-ppt-preview-styles")?.remove();
  FakeIO.instances = [];
  loadMock.mockReset();
  engineRenderMock.mockReset();
  engineDestroyMock.mockReset();
  buildTextLayerMock.mockReset();
  engineRenderMock.mockResolvedValue(undefined);
  loadMock.mockResolvedValue(fakeEngine(3));
});

/** 渲染带 IO（100% margin）。 */
function renderIOs(): FakeIO[] {
  return FakeIO.instances.filter((io) =>
    String(io.options?.rootMargin).startsWith("100%"),
  );
}
function releaseIOs(): FakeIO[] {
  return FakeIO.instances.filter((io) =>
    String(io.options?.rootMargin).startsWith("300%"),
  );
}
function pageAt(i: number): Element {
  return container.querySelectorAll(".mop-ppt-page")[i];
}

describe("renderSlide 参数合同", () => {
  it("固定 mode:'main' 与防御性 buffer；透传 cjkFallback/resourceLimits/workerTimeoutMs/useGoogleFonts 默认关", async () => {
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
      renderOptions: { cjkFallback: "sc", workerTimeoutMs: 1234 },
    });
    await vi.waitFor(() => {
      expect(loadMock).toHaveBeenCalledTimes(1);
    });
    const [buf, opts] = loadMock.mock.calls[0] as [
      ArrayBuffer,
      Record<string, unknown>,
    ];
    expect(buf).toBeInstanceOf(ArrayBuffer);
    expect(opts.mode).toBe("main");
    expect(opts.cjkFallback).toBe("sc");
    expect(opts.workerTimeoutMs).toBe(1234);
    expect(opts.useGoogleFonts).toBe(false);
    expect(opts.password).toBeUndefined();
    inst.destroy();
  });

  it("password 透传；skipMediaControls:true 静态预览；width 按选项（默认 960）", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    loadMock.mockResolvedValue(fakeEngine(1));
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
      password: "pw",
      slideWidth: 640,
      dpr: 2,
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(1);
    });
    expect(loadMock.mock.calls[0][1].password).toBe("pw");
    renderIOs()[0].fire(pageAt(0), true);
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(1);
    });
    const opts = engineRenderMock.mock.calls[0][2] as Record<string, unknown>;
    expect(opts.skipMediaControls).toBe(true);
    expect(opts.width).toBe(640);
    expect(opts.dpr).toBe(2);
    expect(typeof opts.onTextRun).toBe("function");
    inst.destroy();
    vi.unstubAllGlobals();
  });
});

describe("懒渲染队列", () => {
  it("串行渲染：前一页完成后才渲下一页", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const d = deferred();
    let calls = 0;
    engineRenderMock.mockImplementation(() => {
      calls++;
      return calls === 1 ? d.promise : Promise.resolve();
    });
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    const io = renderIOs()[0];
    io.fire(pageAt(0), true);
    io.fire(pageAt(1), true);
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(1);
    });
    d.resolve(); // 第一页完成 → 第二页开始
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(2);
    });
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("离带弃渲：排队中滚出渲染带，出队即弃不渲染", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    // 第一页渲染挂起占住串行队列，让页 2 只能排队（而非立即开渲）
    const d = deferred();
    engineRenderMock.mockImplementationOnce(() => d.promise);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    const io = renderIOs()[0];
    io.fire(pageAt(0), true); // 占住队列
    io.fire(pageAt(2), true); // 入队（排队中）
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(1);
    });
    io.fire(pageAt(2), false); // 排队中滚出 → wanted=false
    d.resolve(); // 队列释放，pump 出队页 2 → 离带即弃
    await new Promise((r) => setTimeout(r, 10));
    expect(engineRenderMock).toHaveBeenCalledTimes(1); // 页 2 未渲
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("释放：滚出保留带后 canvas 置 0 并移除 DOM；重入渲染带可重渲", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    const rio = renderIOs()[0];
    const lio = releaseIOs()[0];
    rio.fire(pageAt(0), true);
    await vi.waitFor(() => {
      expect(pageAt(0).querySelector("canvas")).not.toBeNull();
    });
    const canvas = pageAt(0).querySelector("canvas") as HTMLCanvasElement;
    lio.fire(pageAt(0), false); // 滚出保留带 → 释放
    await vi.waitFor(() => {
      expect(pageAt(0).querySelector("canvas")).toBeNull();
    });
    expect(canvas.width).toBe(0); // 位图已释放
    // 重入渲染带 → 重渲（上游无单页 cancel，释放仅清 canvas/文本 DOM）
    rio.fire(pageAt(0), true);
    await vi.waitFor(() => {
      expect(pageAt(0).querySelector("canvas")).not.toBeNull();
    });
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("在途渲染出保留带：标记过期，完成后释放（引擎无 cancel 句柄的差异语义）", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const d = deferred();
    engineRenderMock.mockReturnValueOnce(d.promise);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    const rio = renderIOs()[0];
    const lio = releaseIOs()[0];
    rio.fire(pageAt(0), true);
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(1);
    });
    lio.fire(pageAt(0), false); // 在途中要求释放
    expect(pageAt(0).querySelector("canvas")).not.toBeNull(); // 还没完成，不假装取消
    d.resolve();
    await vi.waitFor(() => {
      expect(pageAt(0).querySelector("canvas")).toBeNull(); // 完成后收尾释放
    });
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("释放请求后、渲染完成前又回到渲染带：完成时保留本次渲染结果", async () => {
    // 竞态：渲染中滚出保留带（pendingRelease=true），取消落定前用户又滚回
    // 渲染带。renderIO 若在 rendering 状态忽略进入事件，wanted 会停留在
    // false，完成时按过期标记释放刚回到视口的页——页面又变空白。
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const d = deferred();
    engineRenderMock.mockReturnValueOnce(d.promise);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    const rio = renderIOs()[0];
    const lio = releaseIOs()[0];
    rio.fire(pageAt(0), true);
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(1);
    });
    lio.fire(pageAt(0), false); // 在途中要求释放
    rio.fire(pageAt(0), true); // 完成前又回到渲染带
    d.resolve();
    await vi.waitFor(() => {
      expect(pageAt(0).querySelector("canvas")).not.toBeNull();
    });
    // 等本轮 pump 的收尾微任务跑完，确认 canvas 没有被误释放
    await new Promise((r) => setTimeout(r, 20));
    expect(pageAt(0).querySelector("canvas")).not.toBeNull();
    expect(engineRenderMock).toHaveBeenCalledTimes(1); // 保留结果，不重复渲染
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("渲染中 destroy：在途 renderSlide 迟到完成后不写回调、DOM 已清场", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const d = deferred();
    engineRenderMock.mockReturnValueOnce(d.promise); // 首页渲染挂起
    const onRendered = vi.fn();
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
      onRendered,
      onError,
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    renderIOs()[0].fire(pageAt(0), true);
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(1); // 渲染已发出、挂起中
    });
    // 渲染侧先停（canvas 随 renderer.destroy 移除）、引擎随后同步销毁
    inst.destroy();
    expect(engineDestroyMock).toHaveBeenCalledTimes(1);
    d.resolve(); // 迟到完成 → renderEntry 的 destroyed 检查静默
    await new Promise((r) => setTimeout(r, 10));
    expect(onRendered).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
    expect(container.children).toHaveLength(0);
    vi.unstubAllGlobals();
  });
});

describe("onRendered / 错误上报", () => {
  it("首张成功页触发 onRendered 一次（第 0 页失败不阻塞、不触发）", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    engineRenderMock.mockImplementation((_c: unknown, i: number) =>
      i === 0 ? Promise.reject(new Error("slide 0 broken")) : Promise.resolve(),
    );
    const onRendered = vi.fn();
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
      onRendered,
      onError,
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    const rio = renderIOs()[0];
    rio.fire(pageAt(0), true);
    await vi.waitFor(() => {
      expect(onError).toHaveBeenCalledTimes(1); // 第 0 页失败上报
    });
    rio.fire(pageAt(1), true);
    await vi.waitFor(() => {
      expect(onRendered).toHaveBeenCalledTimes(1); // 首张成功页（页 1）触发
    });
    const info = onRendered.mock.calls[0][0];
    expect(info.slideCount).toBe(3);
    expect(info.firstSlideMs).toBeGreaterThanOrEqual(0);
    expect(info.durationMs).toBeGreaterThanOrEqual(0);
    rio.fire(pageAt(2), true);
    await new Promise((r) => setTimeout(r, 5));
    expect(onRendered).toHaveBeenCalledTimes(1); // 只一次
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("渲染失败经 onError 上报归一错误（引擎错误码 → 公开码）", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    engineRenderMock.mockRejectedValueOnce(
      Object.assign(new Error("declared-inflated-bytes over"), {
        code: "ooxml-resource-limit",
      }),
    );
    const onError = vi.fn();
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
      onError,
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    renderIOs()[0].fire(pageAt(0), true);
    await vi.waitFor(() => {
      expect(onError).toHaveBeenCalledTimes(1);
    });
    expect(onError.mock.calls[0][0].code).toBe("UNSUPPORTED");
    inst.destroy();
    vi.unstubAllGlobals();
  });
});

describe("0-based 导航（含隐藏页的原始页序）", () => {
  it("goToSlide 非整数 → TypeError；加载前 → no-op", () => {
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    expect(() => inst.goToSlide(1.5)).toThrow(TypeError);
    expect(() => inst.goToSlide(Number.NaN)).toThrow(TypeError);
    inst.goToSlide(0); // 加载前 no-op，不抛
    inst.destroy();
  });

  it("加载后越界 → RangeError；合法索引 no-op 不抛", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    expect(() => inst.goToSlide(3)).toThrow(RangeError);
    expect(() => inst.goToSlide(-1)).toThrow(RangeError);
    expect(() => inst.goToSlide(2)).not.toThrow();
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("slideCount：加载前 0、加载后为原始总页数、destroy 后 0", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    expect(inst.slideCount).toBe(0);
    await vi.waitFor(() => {
      expect(inst.slideCount).toBe(3);
    });
    inst.destroy();
    expect(inst.slideCount).toBe(0);
    vi.unstubAllGlobals();
  });

  it("next/prev 基于当前最可见页滚动、goToSlide 直达；末页边界 no-op", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    // happy-dom 无布局：rect 全 0 → 最可见页为 -1 → next/prev no-op 不抛
    expect(() => {
      inst.nextSlide();
      inst.prevSlide();
    }).not.toThrow();

    // 注入可控 rect + spy scrollBy。页高远小于 happy-dom 默认视口（768px），
    // 交叠面积即页高（结论不随 innerHeight 波动）：页 1（y 50-350，交叠
    // 300）独占最大 → 最可见。scrollRoot 为 null（无滚动祖先）→ 走
    // window.scrollBy，参数为目标页 rect.top
    const rects = new Map<Element, DOMRect>();
    rects.set(pageAt(0), new DOMRect(0, 0, 960, 50));
    rects.set(pageAt(1), new DOMRect(0, 50, 960, 300));
    rects.set(pageAt(2), new DOMRect(0, 350, 960, 50));
    const rectSpy = vi
      .spyOn(Element.prototype, "getBoundingClientRect")
      .mockImplementation(function (this: Element) {
        return rects.get(this) ?? new DOMRect();
      });
    const scrollBySpy = vi
      .spyOn(window, "scrollBy")
      .mockImplementation(() => {});

    inst.nextSlide(); // 最可见页 1 → 滚到页 2（rect.top = 350）
    expect(scrollBySpy).toHaveBeenCalledWith({ top: 350 });
    scrollBySpy.mockClear();
    inst.prevSlide(); // 最可见页 1 → 滚到页 0（rect.top = 0）
    expect(scrollBySpy).toHaveBeenCalledWith({ top: 0 });
    scrollBySpy.mockClear();
    inst.goToSlide(2); // 指定索引直达（rect.top = 350）
    expect(scrollBySpy).toHaveBeenCalledWith({ top: 350 });
    scrollBySpy.mockClear();

    // 最可见页 = 末页时 nextSlide 边界 no-op（不触发滚动）
    rects.set(pageAt(1), new DOMRect(0, 50, 960, 50));
    rects.set(pageAt(2), new DOMRect(0, 350, 960, 300));
    inst.nextSlide();
    expect(scrollBySpy).not.toHaveBeenCalled();

    scrollBySpy.mockRestore();
    rectSpy.mockRestore();
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("滚动容器路径：定高 overflow:auto 面板成为 IO root，导航滚动容器而非窗口", async () => {
    // quick-start 宣称的主用法：定高 overflow:auto 面板。inline 样式即可
    // 驱动 happy-dom 的 computed style 与 hasDefiniteHeight 判定
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const panel = document.createElement("div");
    panel.style.overflowY = "auto";
    panel.style.height = "400px";
    panel.appendChild(container); // beforeEach 已把 container 挂 body，此处移入面板
    document.body.appendChild(panel);

    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    // findScrollRoot 识别：双 IO 的 root 绑定到面板（而非视口 null）
    expect(renderIOs()[0].options?.root).toBe(panel);
    expect(releaseIOs()[0].options?.root).toBe(panel);

    // 注入 rect：面板视口 y 0-400；页 0 在 0-100、页 1 在 100-300（交叠
    // 最大 → 最可见）、页 2 在 700-800
    const rects = new Map<Element, DOMRect>();
    rects.set(panel, new DOMRect(0, 0, 960, 400));
    rects.set(pageAt(0), new DOMRect(0, 0, 960, 100));
    rects.set(pageAt(1), new DOMRect(0, 100, 960, 200));
    rects.set(pageAt(2), new DOMRect(0, 700, 960, 100));
    const rectSpy = vi
      .spyOn(Element.prototype, "getBoundingClientRect")
      .mockImplementation(function (this: Element) {
        return rects.get(this) ?? new DOMRect();
      });
    const scrollBySpy = vi
      .spyOn(window, "scrollBy")
      .mockImplementation(() => {});

    // 最可见页按面板视口计算：页 1 → next 滚到页 2（top 700 - 面板顶 0）
    panel.scrollTop = 0;
    inst.nextSlide();
    expect(panel.scrollTop).toBe(700);
    expect(scrollBySpy).not.toHaveBeenCalled(); // 不连带滚动窗口

    // goToSlide 相对修正：从 scrollTop=10 出发滚到页 1 → 10 + (100-0)
    panel.scrollTop = 10;
    inst.goToSlide(1);
    expect(panel.scrollTop).toBe(110);

    scrollBySpy.mockRestore();
    rectSpy.mockRestore();
    inst.destroy();
    panel.remove();
    vi.unstubAllGlobals();
  });

  it("destroy 后导航全部 no-op", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(3);
    });
    inst.destroy();
    expect(() => inst.goToSlide(0)).not.toThrow(); // destroy 后 no-op（不抛 RangeError）
    inst.nextSlide();
    inst.prevSlide();
    vi.unstubAllGlobals();
  });
});
