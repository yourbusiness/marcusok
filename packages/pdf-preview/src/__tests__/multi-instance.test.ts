// @vitest-environment happy-dom
/**
 * 多实例测试：双容器双实例互不串扰（各自的懒渲染管线独立工作），
 * 文本层样式单例只注入一次（文档级共享，不随 destroy 删除）。
 */
import { describe, it, expect, vi, beforeEach } from "vitest";
import { createPdfPreview } from "../render/render";

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

const getDocumentMock = vi.hoisted(() => vi.fn());
// 模拟 pdfjs 真实形态：cleanup 是静态方法且内部读 this（#pendingTextLayers），
// 接收者丢失（裸调用）会抛 TypeError——用于锁定 releaseGlobalTextLayerCaches
// 必须以 TextLayer 自身为接收者调用 cleanup（此前裸调用在浏览器实测必抛、
// 被包层 catch 降级，全局文本层缓存清理从未生效）。
const textLayerCleanupMock = vi.hoisted(() =>
  vi.fn(function (this: unknown) {
    if (this === undefined || this === null) {
      throw new TypeError(
        "Cannot read properties of undefined (reading '#pendingTextLayers')",
      );
    }
  }),
);
const TextLayerMock = vi.hoisted(() =>
  vi.fn(function () {
    return { render: vi.fn(async () => {}), cancel: vi.fn() };
  }),
);
// vi.fn 的推断类型不带静态位，静态属性经 Object.assign 挂上（模拟 pdfjs
// 的 class TextLayer { static cleanup }）
Object.assign(TextLayerMock, { cleanup: textLayerCleanupMock });
vi.mock("pdfjs-dist", () => ({
  getDocument: getDocumentMock,
  GlobalWorkerOptions: { workerSrc: "" },
  TextLayer: TextLayerMock,
  InvalidPDFException: class extends Error {},
  PasswordException: class extends Error {},
  RenderingCancelledException: class extends Error {},
}));

function fakePage(num: number, width = 612) {
  return {
    pageNumber: num,
    getViewport: ({ scale = 1 }: { scale?: number }) => ({
      width: width * scale,
      height: 792 * scale,
      scale,
    }),
    render: vi.fn(() => ({ promise: Promise.resolve(), cancel: vi.fn() })),
    cleanup: vi.fn(() => true),
    streamTextContent: () => ({}),
  };
}

function magicSource(): Uint8Array {
  const bytes = new Uint8Array(64);
  bytes.set([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x34], 0);
  return bytes;
}

let a: HTMLElement;
let b: HTMLElement;

beforeEach(() => {
  document.body.textContent = "";
  a = document.createElement("div");
  b = document.createElement("div");
  document.body.append(a, b);
  document.getElementById("marcusok-pdf-preview-styles")?.remove();
  FakeIO.instances = [];
  getDocumentMock.mockReset();
  textLayerCleanupMock.mockClear();
  getDocumentMock.mockImplementation(() => ({
    promise: Promise.resolve({
      numPages: 1,
      getPage: () => Promise.resolve(fakePage(1)),
    }),
    destroy: vi.fn(async () => {}),
  }));
});

describe("multi-instance isolation", () => {
  it("injects the shared stylesheet exactly once across instances", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const i1 = createPdfPreview(a, { source: magicSource() });
    const i2 = createPdfPreview(b, { source: magicSource() });
    await vi.waitFor(() => {
      expect(a.querySelectorAll(".mop-pdf-page")).toHaveLength(1);
      expect(b.querySelectorAll(".mop-pdf-page")).toHaveLength(1);
    });
    expect(
      document.querySelectorAll("#marcusok-pdf-preview-styles"),
    ).toHaveLength(1);
    i1.destroy();
    i2.destroy();
    // 样式单例不随 destroy 删除（多实例共享）
    expect(
      document.getElementById("marcusok-pdf-preview-styles"),
    ).not.toBeNull();
  });

  it("renders independently per container and destroy of one leaves the other working", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const i1 = createPdfPreview(a, { source: magicSource() });
    const i2 = createPdfPreview(b, { source: magicSource() });
    await vi.waitFor(() => {
      expect(a.querySelectorAll(".mop-pdf-page")).toHaveLength(1);
      expect(b.querySelectorAll(".mop-pdf-page")).toHaveLength(1);
    });

    // 每实例各有自己的 renderIO（FakeIO.instances 顺序：i1 render/release, i2 render/release）
    const io100 = FakeIO.instances.filter((io) =>
      String(io.options?.rootMargin).startsWith("100%"),
    );
    expect(io100).toHaveLength(2);
    io100[0].fire(a.querySelector(".mop-pdf-page") as Element, true);
    await vi.waitFor(() => {
      expect(a.querySelector(".mop-pdf-page canvas")).not.toBeNull();
    });
    expect(b.querySelector(".mop-pdf-page canvas")).toBeNull(); // 未触发不渲

    i1.destroy();
    expect(a.children).toHaveLength(0);
    expect(b.querySelectorAll(".mop-pdf-page")).toHaveLength(1); // i2 不受影响

    io100[1].fire(b.querySelector(".mop-pdf-page") as Element, true);
    await vi.waitFor(() => {
      expect(b.querySelector(".mop-pdf-page canvas")).not.toBeNull();
    });
    i2.destroy();
  });

  it("最后实例销毁时以 TextLayer 为接收者调用静态 cleanup", () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const i1 = createPdfPreview(a, { source: magicSource() });
    const i2 = createPdfPreview(b, { source: magicSource() });
    // 双实例在世：cleanup 不得提前触发（活跃计数守护）
    i1.destroy();
    expect(textLayerCleanupMock).not.toHaveBeenCalled();
    i2.destroy();
    expect(textLayerCleanupMock).toHaveBeenCalledTimes(1);
    // 静态方法依赖 this：裸调用（接收者 undefined）在真实 pdfjs 里抛
    // TypeError。contexts 记录每次调用的 this——必须等于 TextLayer 本身。
    expect(textLayerCleanupMock.mock.contexts[0]).toBe(TextLayerMock);
  });
});
