// @vitest-environment happy-dom
/**
 * 多实例隔离：双容器双实例各自懒渲染互不串扰、样式单例只注入一次、
 * 一个实例 destroy 不影响另一个。
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const loadMock = vi.hoisted(() => vi.fn());
const engineRenderMock = vi.hoisted(() => vi.fn());
const engineDestroyMock = vi.hoisted(() => vi.fn());
vi.mock("@silurus/ooxml/pptx", () => ({
  PptxPresentation: { load: (...a: unknown[]) => loadMock(...a) },
  buildPptxTextLayer: vi.fn(),
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

function fakeEngine(slides: number, marker: string) {
  return {
    slideCount: slides,
    slideWidth: 12192000,
    slideHeight: 6858000,
    renderSlide: (canvas: HTMLCanvasElement, i: number) =>
      engineRenderMock(marker, canvas, i),
    destroy: () => engineDestroyMock(marker),
  };
}

let a: HTMLElement;
let b: HTMLElement;

beforeEach(() => {
  document.body.textContent = "";
  a = document.createElement("div");
  b = document.createElement("div");
  document.body.append(a, b);
  document.getElementById("marcusok-ppt-preview-styles")?.remove();
  FakeIO.instances = [];
  loadMock.mockReset();
  engineRenderMock.mockReset();
  engineDestroyMock.mockReset();
  engineRenderMock.mockResolvedValue(undefined);
  // 默认引擎（单页）；需要不同页数的测试自行 mockImplementationOnce 覆盖
  loadMock.mockResolvedValue(fakeEngine(1, "default"));
});

describe("多实例隔离", () => {
  it("双实例独立加载与懒渲染；destroy 一个不影响另一个", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    loadMock.mockImplementationOnce(() => Promise.resolve(fakeEngine(2, "A")));
    loadMock.mockImplementationOnce(() => Promise.resolve(fakeEngine(5, "B")));
    const i1 = createPptPreview(a, { source: buildPptx({ texts: ["a"] }) });
    const i2 = createPptPreview(b, { source: buildPptx({ texts: ["a"] }) });
    await vi.waitFor(() => {
      expect(a.querySelectorAll(".mop-ppt-page")).toHaveLength(2);
      expect(b.querySelectorAll(".mop-ppt-page")).toHaveLength(5);
    });
    // 各实例的渲染带 IO 独立（100% margin 的前两个）
    const ioA = FakeIO.instances.filter((io) =>
      String(io.options?.rootMargin).startsWith("100%"),
    )[0];
    const ioB = FakeIO.instances.filter((io) =>
      String(io.options?.rootMargin).startsWith("100%"),
    )[1];
    ioA.fire(a.querySelector(".mop-ppt-page") as Element, true);
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(1);
    });
    expect((engineRenderMock.mock.calls[0] as unknown[])[0]).toBe("A");
    expect(b.querySelector("canvas")).toBeNull(); // 未触发不渲
    ioB.fire(b.querySelector(".mop-ppt-page") as Element, true);
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(2);
    });
    expect((engineRenderMock.mock.calls[1] as unknown[])[0]).toBe("B");

    i1.destroy();
    expect(engineDestroyMock).toHaveBeenCalledWith("A");
    expect(a.children).toHaveLength(0);
    expect(b.querySelectorAll(".mop-ppt-page")).toHaveLength(5); // B 不受影响
    i2.destroy();
    expect(engineDestroyMock).toHaveBeenCalledWith("B");
    vi.unstubAllGlobals();
  });

  it("样式表只注入一次（文档级共享，不随 destroy 删除）", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const i1 = createPptPreview(a, { source: buildPptx({ texts: ["a"] }) });
    const i2 = createPptPreview(b, { source: buildPptx({ texts: ["a"] }) });
    await vi.waitFor(() => {
      expect(a.querySelectorAll(".mop-ppt-page")).toHaveLength(1);
    });
    expect(
      document.querySelectorAll("#marcusok-ppt-preview-styles"),
    ).toHaveLength(1);
    i1.destroy();
    i2.destroy();
    expect(
      document.getElementById("marcusok-ppt-preview-styles"),
    ).not.toBeNull();
    vi.unstubAllGlobals();
  });

  it("slideCount 各读各的，destroy 只清自己", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    loadMock.mockImplementationOnce(() => Promise.resolve(fakeEngine(2, "A")));
    loadMock.mockImplementationOnce(() => Promise.resolve(fakeEngine(5, "B")));
    const i1 = createPptPreview(a, { source: buildPptx({ texts: ["a"] }) });
    const i2 = createPptPreview(b, { source: buildPptx({ texts: ["a"] }) });
    await vi.waitFor(() => {
      expect(i1.slideCount).toBe(2);
      expect(i2.slideCount).toBe(5);
    });
    i1.destroy();
    expect(i1.slideCount).toBe(0);
    expect(i2.slideCount).toBe(5);
    vi.unstubAllGlobals();
  });
});
