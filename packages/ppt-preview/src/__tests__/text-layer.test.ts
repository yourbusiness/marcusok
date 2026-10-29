// @vitest-environment happy-dom
/**
 * 文本层测试：单趟收集（onTextRun 与正式渲染同趟、不调用
 * collectSlideRuns 的二次渲染路径）、buildPptxTextLayer 的调用合同
 * （不传 click 回调 → 链接 inert）、enableTextSelection:false 关闭。
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const loadMock = vi.hoisted(() => vi.fn());
const engineRenderMock = vi.hoisted(() => vi.fn());
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
  constructor(cb: IntersectionObserverCallback) {
    this.callback = cb;
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

function fakeEngine(slides = 1) {
  return {
    slideCount: slides,
    slideWidth: 12192000,
    slideHeight: 6858000,
    renderSlide: (...a: unknown[]) => engineRenderMock(...a),
    destroy: vi.fn(),
  };
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
  buildTextLayerMock.mockReset();
  engineRenderMock.mockImplementation(
    (
      _canvas: unknown,
      _i: unknown,
      opts: { onTextRun?: (r: unknown) => void },
    ) => {
      // 模拟上游：渲染中同步回调 runs（真实上游行为，spike 实测 3 runs）
      opts.onTextRun?.({
        text: "标题文本",
        inShapeX: 1,
        inShapeY: 2,
        w: 10,
        h: 5,
      });
      return Promise.resolve();
    },
  );
  loadMock.mockResolvedValue(fakeEngine(1));
});

describe("文本层单趟收集", () => {
  it("onTextRun 与正式渲染同趟；runs 原样传给 buildPptxTextLayer；不传 click 回调（链接 inert）", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(1);
    });
    FakeIO.instances[0].fire(
      container.querySelector(".mop-ppt-page") as Element,
      true,
    );
    await vi.waitFor(() => {
      expect(buildTextLayerMock).toHaveBeenCalledTimes(1);
    });
    const [layerEl, runs, cssW, cssH, clickCb] =
      buildTextLayerMock.mock.calls[0];
    expect(layerEl).toBeInstanceOf(HTMLDivElement);
    expect(layerEl.className).toContain("mop-ppt-text-layer");
    expect(runs).toHaveLength(1);
    expect((runs[0] as { text: string }).text).toBe("标题文本");
    expect(cssW).toBe(960); // 默认 slideWidth
    expect(cssH).toBe(540); // 16:9
    expect(clickCb).toBeUndefined(); // 链接 inert 的合同：不传 onHyperlinkClick
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("enableTextSelection:false 不传 onTextRun、不建文本层", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
      enableTextSelection: false,
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(1);
    });
    FakeIO.instances[0].fire(
      container.querySelector(".mop-ppt-page") as Element,
      true,
    );
    await vi.waitFor(() => {
      expect(engineRenderMock).toHaveBeenCalledTimes(1);
    });
    const opts = engineRenderMock.mock.calls[0][2] as { onTextRun?: unknown };
    expect(opts.onTextRun).toBeUndefined();
    await new Promise((r) => setTimeout(r, 5));
    expect(buildTextLayerMock).not.toHaveBeenCalled();
    expect(container.querySelector(".mop-ppt-text-layer")).toBeNull();
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("释放时文本层随占位清理", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPptPreview(container, {
      source: buildPptx({ texts: ["a"] }),
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-ppt-page")).toHaveLength(1);
    });
    FakeIO.instances[0].fire(
      container.querySelector(".mop-ppt-page") as Element,
      true,
    );
    await vi.waitFor(() => {
      expect(container.querySelector(".mop-ppt-text-layer")).not.toBeNull();
    });
    inst.destroy();
    expect(container.querySelector(".mop-ppt-text-layer")).toBeNull();
    vi.unstubAllGlobals();
  });

  it("styleNonce 只作用于本包 <style>（幂等单例）", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const a = document.createElement("div");
    const b = document.createElement("div");
    document.body.append(a, b);
    const i1 = createPptPreview(a, {
      source: buildPptx({ texts: ["a"] }),
      styleNonce: "n1",
    });
    const i2 = createPptPreview(b, {
      source: buildPptx({ texts: ["a"] }),
      styleNonce: "n2",
    });
    await vi.waitFor(() => {
      expect(a.querySelectorAll(".mop-ppt-page")).toHaveLength(1);
    });
    const styles = document.querySelectorAll("#marcusok-ppt-preview-styles");
    expect(styles).toHaveLength(1); // 单例注入
    expect((styles[0] as HTMLElement).getAttribute("nonce")).toBe("n1"); // 先到先得
    i1.destroy();
    i2.destroy();
    vi.unstubAllGlobals();
  });
});
