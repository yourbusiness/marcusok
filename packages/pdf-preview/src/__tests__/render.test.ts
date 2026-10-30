// @vitest-environment happy-dom
/**
 * render 层测试：懒渲染状态机、缩放解析、串行队列、释放/重渲、错误归一。
 *
 * 环境边界（与 word-preview render.test 同口径）：
 * - happy-dom 的 canvas 无 2d context、IntersectionObserver 无真实交叉
 *   计算——本文件 mock pdfjs-dist（getDocument/TextLayer 返回可控假件）+
 *   FakeIO 手动派发回调，只测**本包自己的调度逻辑**；
 * - 真实解析在 parse.test（node + legacy 引擎）覆盖、真实渲染保真在
 *   play/docs demo 的浏览器验收覆盖，此处不声称渲染保真。
 */
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { createPdfPreview } from "../render/render";
import type { PdfPreviewError } from "../types";

/* ------------------------------- Fake IO ------------------------------- */

class FakeIO {
  static instances: FakeIO[] = [];
  callback: IntersectionObserverCallback;
  options: IntersectionObserverInit | undefined;
  disconnected = false;
  private targets = new Set<Element>();

  constructor(
    cb: IntersectionObserverCallback,
    options?: IntersectionObserverInit,
  ) {
    this.callback = cb;
    this.options = options;
    FakeIO.instances.push(this);
  }
  observe(el: Element) {
    this.targets.add(el);
  }
  unobserve(el: Element) {
    this.targets.delete(el);
  }
  disconnect() {
    this.disconnected = true;
    this.targets.clear();
  }
  fire(el: Element, isIntersecting: boolean) {
    this.callback(
      [{ target: el, isIntersecting } as IntersectionObserverEntry],
      this as unknown as IntersectionObserver,
    );
  }
}

/** 取实例：margin "100%" 为渲染观察器、"300%" 为释放观察器。 */
function ioByMargin(marginPrefix: string): FakeIO[] {
  return FakeIO.instances.filter((io) =>
    String(io.options?.rootMargin ?? "").startsWith(marginPrefix),
  );
}

/* ---------------------------- mocked pdfjs ---------------------------- */

// vi.mock 工厂被提升到文件顶部，其引用的值必须同为 vi.hoisted 产物
const FakePasswordException = vi.hoisted(
  () =>
    class FakePasswordException extends Error {
      code: number;
      constructor(msg: string, code: number) {
        super(msg);
        this.name = "PasswordException";
        this.code = code;
      }
    },
);

interface Deferred<T> {
  promise: Promise<T>;
  resolve: (v: T) => void;
  reject: (e: unknown) => void;
}
function deferred<T>(): Deferred<T> {
  let resolve!: (v: T) => void;
  let reject!: (e: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

export interface FakePageOptions {
  width?: number;
  height?: number;
  /** render 返回受控 deferred（串行队列测试用）；缺省立即成功。 */
  deferred?: boolean;
  failRender?: Error;
  /** deferred 模式下 cancel() 的拒绝错误名（默认 "Error"，真实引擎为 RenderingCancelledException）。 */
  cancelName?: string;
}

/** 假 PDFPageProxy：默认 612×792（US Letter），记录 render 调用参数。 */
function fakePage(num: number, opts: FakePageOptions = {}) {
  const { width = 612, height = 792 } = opts;
  const renderCalls: Array<{
    canvas: HTMLCanvasElement;
    scale: number;
    transform?: number[];
  }> = [];
  /** deferred 模式下每次 render 的放行句柄（串行队列测试用）。 */
  const renderDeferreds: Deferred<void>[] = [];
  const page = {
    pageNumber: num,
    getViewport: ({ scale = 1 }: { scale?: number }) => ({
      width: width * scale,
      height: height * scale,
      scale,
    }),
    render: vi.fn(
      (params: {
        canvas: HTMLCanvasElement;
        viewport: { scale: number };
        transform?: number[];
      }) => {
        renderCalls.push({
          canvas: params.canvas,
          scale: params.viewport.scale,
          transform: params.transform,
        });
        if (opts.failRender) {
          return { promise: Promise.reject(opts.failRender), cancel: vi.fn() };
        }
        if (opts.deferred) {
          const d = deferred<void>();
          renderDeferreds.push(d);
          return {
            promise: d.promise,
            cancel: vi.fn(() => {
              const err = new Error("cancel");
              if (opts.cancelName) err.name = opts.cancelName;
              d.reject(err);
            }),
          };
        }
        return { promise: Promise.resolve(), cancel: vi.fn() };
      },
    ),
    renderCalls,
    renderDeferreds,
    cleanup: vi.fn(() => true),
    streamTextContent: () => ({ fake: "stream" }),
  };
  return page;
}

type FakePage = ReturnType<typeof fakePage>;

function fakeDoc(pages: FakePage[], opts: { pendingGetPage?: boolean } = {}) {
  const getPageDeferred = deferred<void>();
  return {
    doc: {
      numPages: pages.length,
      getPage: (n: number) =>
        opts.pendingGetPage
          ? getPageDeferred.promise.then(() => pages[n - 1])
          : Promise.resolve(pages[n - 1]),
    },
    getPageDeferred,
  };
}

const getDocumentMock = vi.hoisted(() => vi.fn());
const GlobalWorkerOptionsMock = vi.hoisted(() => ({ workerSrc: "" }));
// 注意实现必须是普通 function（可 new）：箭头函数实现的 vi.fn 在
// `new TextLayer()` 时抛 "not a constructor"，会把渲染错误路径全带偏
const TextLayerMock = vi.hoisted(() =>
  vi.fn(function fakeTextLayer() {
    return { render: vi.fn(async () => {}), cancel: vi.fn() };
  }),
);

vi.mock("pdfjs-dist", () => ({
  getDocument: getDocumentMock,
  GlobalWorkerOptions: GlobalWorkerOptionsMock,
  TextLayer: TextLayerMock,
  InvalidPDFException: class extends Error {
    constructor(msg: string) {
      super(msg);
      this.name = "InvalidPDFException";
    }
  },
  PasswordException: FakePasswordException,
  RenderingCancelledException: class extends Error {
    constructor(msg: string) {
      super(msg);
      this.name = "RenderingCancelledException";
    }
  },
}));

import { getDocument } from "pdfjs-dist";

/** 嗅探可过、getDocument 走 mock 的最小输入源。 */
function magicSource(): Uint8Array {
  const bytes = new Uint8Array(64);
  bytes.set([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x34], 0);
  return bytes;
}

function setupDocument(doc: object, loadingTaskOverride?: object) {
  getDocumentMock.mockImplementation(() => ({
    promise: Promise.resolve(doc),
    destroy: vi.fn(async () => {}),
    ...loadingTaskOverride,
  }));
}

function setClientSize(el: HTMLElement, width: number, height: number) {
  Object.defineProperty(el, "clientWidth", {
    value: width,
    configurable: true,
  });
  Object.defineProperty(el, "clientHeight", {
    value: height,
    configurable: true,
  });
}

async function waitForPages(container: HTMLElement, n: number) {
  await vi.waitFor(() => {
    expect(container.querySelectorAll(".mop-pdf-page")).toHaveLength(n);
  });
}

const pageEls = (container: HTMLElement) =>
  Array.from(container.querySelectorAll<HTMLElement>(".mop-pdf-page"));

let container: HTMLElement;

beforeEach(() => {
  document.body.textContent = "";
  container = document.createElement("div");
  document.body.appendChild(container);
  // 样式单例是文档级的：跨测试残留会让 styleNonce 断言失真，逐测试清掉
  document.getElementById("marcusok-pdf-preview-styles")?.remove();
  FakeIO.instances = [];
  getDocumentMock.mockReset();
  TextLayerMock.mockClear();
});

afterEach(() => {
  vi.unstubAllGlobals();
});

/* ------------------------------- tests ------------------------------- */

describe("createPdfPreview render pipeline (mocked engine)", () => {
  it("creates per-page placeholders sized by viewport*scale (numeric scale)", async () => {
    const { doc } = fakeDoc([fakePage(1), fakePage(2)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, {
      source: magicSource(),
      scale: 2,
    });
    await waitForPages(container, 2);
    const els = pageEls(container);
    expect(els[0].style.width).toBe("1224px");
    expect(els[0].style.height).toBe("1584px");
    // 页级 CSS 变量（文本层字号计算链输入）
    expect(els[0].style.getPropertyValue("--scale-factor")).toBe("2");
    inst.destroy();
  });

  it("fit-width resolves scale from container width / first page width", async () => {
    const { doc } = fakeDoc([fakePage(1), fakePage(2)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    setClientSize(container, 306, 1000); // 306/612 = 0.5
    const inst = createPdfPreview(container, { source: magicSource() });
    await waitForPages(container, 2);
    expect(pageEls(container)[0].style.width).toBe("306px");
    inst.destroy();
  });

  it("fit-page takes the smaller width/height ratio", async () => {
    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    // 400/612≈0.65 vs 500/792≈0.63 → 取 0.631...
    setClientSize(container, 400, 500);
    const inst = createPdfPreview(container, {
      source: magicSource(),
      scale: "fit-page",
    });
    await waitForPages(container, 1);
    const w = Number.parseFloat(pageEls(container)[0].style.width);
    expect(w).toBeGreaterThan(385);
    expect(w).toBeLessThan(395);
    inst.destroy();
  });

  it("zero-width container falls back to scale=1 with a warning", async () => {
    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    setClientSize(container, 0, 0);
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const inst = createPdfPreview(container, { source: magicSource() });
    await waitForPages(container, 1);
    expect(pageEls(container)[0].style.width).toBe("612px");
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();
    inst.destroy();
  });

  it("renders on render-IO intersect: canvas + textLayer appended, onRendered fired once", async () => {
    const pages = [fakePage(1), fakePage(2)];
    const { doc } = fakeDoc(pages);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onRendered = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onRendered,
    });
    await waitForPages(container, 2);
    const els = pageEls(container);

    ioByMargin("100%")[0].fire(els[0], true);
    await vi.waitFor(() => {
      expect(els[0].querySelector("canvas")).not.toBeNull();
    });
    await vi.waitFor(() => {
      expect(els[0].querySelector(".textLayer")).not.toBeNull();
    });
    expect(els[0].querySelector("canvas")!.style.width).toBe("612px");

    expect(onRendered).toHaveBeenCalledTimes(1);
    const info = onRendered.mock.calls[0][0];
    expect(info.pageCount).toBe(2);
    expect(info.firstPageMs).toBeGreaterThanOrEqual(0);
    expect(info.durationMs).toBeGreaterThanOrEqual(0);

    // 第二页渲染不重复触发 onRendered
    ioByMargin("100%")[0].fire(els[1], true);
    await vi.waitFor(() => {
      expect(pages[1].render).toHaveBeenCalled();
    });
    expect(onRendered).toHaveBeenCalledTimes(1);
    inst.destroy();
  });

  it("textLayer:false renders canvas only", async () => {
    const pages = [fakePage(1)];
    const { doc } = fakeDoc(pages);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, {
      source: magicSource(),
      textLayer: false,
    });
    await waitForPages(container, 1);
    ioByMargin("100%")[0].fire(pageEls(container)[0], true);
    await vi.waitFor(() => {
      expect(pageEls(container)[0].querySelector("canvas")).not.toBeNull();
    });
    expect(pageEls(container)[0].querySelector(".textLayer")).toBeNull();
    expect(TextLayerMock).not.toHaveBeenCalled();
    inst.destroy();
  });

  it("passes viewport scale to render and dpr transform when devicePixelRatio > 1", async () => {
    const pages = [fakePage(1)];
    const { doc } = fakeDoc(pages);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    vi.stubGlobal("devicePixelRatio", 2);
    const inst = createPdfPreview(container, {
      source: magicSource(),
      scale: 1.5,
    });
    await waitForPages(container, 1);
    ioByMargin("100%")[0].fire(pageEls(container)[0], true);
    await vi.waitFor(() => {
      expect(pages[0].render).toHaveBeenCalled();
    });
    const call = pages[0].renderCalls[0];
    expect(call.scale).toBe(1.5);
    expect(call.transform).toEqual([2, 0, 0, 2, 0, 0]);
    const canvas = pageEls(container)[0].querySelector("canvas")!;
    // 位图 = CSS 尺寸 × dpr
    expect(canvas.width).toBe(Math.floor(612 * 1.5 * 2));
    inst.destroy();
  });

  it("serializes renders: the second page starts only after the first resolves", async () => {
    const pages = [
      fakePage(1, { deferred: true }),
      fakePage(2, { deferred: true }),
    ];
    const { doc } = fakeDoc(pages);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, { source: magicSource() });
    await waitForPages(container, 2);
    const els = pageEls(container);

    ioByMargin("100%")[0].fire(els[0], true);
    ioByMargin("100%")[0].fire(els[1], true);
    await vi.waitFor(() => expect(pages[0].render).toHaveBeenCalled());
    // 串行：首页 render 挂起期间，第二页不得开渲
    expect(pages[1].render).not.toHaveBeenCalled();

    pages[0].renderDeferreds[0].resolve(); // 放行首页
    await vi.waitFor(() => expect(pages[1].render).toHaveBeenCalled());
    pages[1].renderDeferreds[0].resolve();
    await vi.waitFor(() => {
      expect(els[1].querySelector("canvas")).not.toBeNull();
    });
    inst.destroy();
  });

  it("releases the canvas when the page leaves the release band and re-renders on return", async () => {
    const pages = [fakePage(1)];
    const { doc } = fakeDoc(pages);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, { source: magicSource() });
    await waitForPages(container, 1);
    const el = pageEls(container)[0];

    ioByMargin("100%")[0].fire(el, true);
    await vi.waitFor(() => {
      expect(el.querySelector("canvas")).not.toBeNull();
    });

    // 离开渲染带（wanted=false）+ 离开保留带（释放）。释放可能经
    // pendingRelease 异步收尾（渲染在途时先记标记、完成后释放），用 waitFor
    ioByMargin("100%")[0].fire(el, false);
    ioByMargin("300%")[0].fire(el, false);
    await vi.waitFor(() => {
      expect(el.querySelector("canvas")).toBeNull();
    });
    expect(el.querySelector(".textLayer")).toBeNull();

    // 重新进入：按同路径重渲
    ioByMargin("100%")[0].fire(el, true);
    await vi.waitFor(() => {
      expect(el.querySelector("canvas")).not.toBeNull();
    });
    expect(pages[0].render).toHaveBeenCalledTimes(2);
    inst.destroy();
  });

  it("re-renders after a cancelled in-flight render when the page re-enters the band", async () => {
    // 竞态：页面在渲中被滚出（release 取消任务），取消尚未落定前用户又滚回。
    // renderIO 的回调若无视 rendering 状态，wanted 会停留在 false，取消完成后
    // 页面回到 empty 却无人重新入队——永远空白，直到再滚出/滚入一次。
    const page = fakePage(1);
    const d = deferred<void>();
    const render = vi
      .fn()
      .mockImplementationOnce(() => ({
        promise: d.promise,
        cancel: vi.fn(() => {
          d.reject(
            Object.assign(new Error("cancelled"), {
              name: "RenderingCancelledException",
            }),
          );
        }),
      }))
      .mockImplementation(() => ({
        promise: Promise.resolve(),
        cancel: vi.fn(),
      }));
    (page as unknown as { render: unknown }).render = render;
    const { doc } = fakeDoc([page]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onError,
    });
    await waitForPages(container, 1);
    const el = pageEls(container)[0];

    ioByMargin("100%")[0].fire(el, true);
    await vi.waitFor(() => expect(render).toHaveBeenCalledTimes(1));

    // 滚出渲染带 + 滚出保留带 → 取消在途渲染；随后立刻又进入渲染带
    ioByMargin("100%")[0].fire(el, false);
    ioByMargin("300%")[0].fire(el, false);
    ioByMargin("100%")[0].fire(el, true);

    // 取消落定后必须重新入队渲染（第二次调用 resolve）
    await vi.waitFor(() => expect(render).toHaveBeenCalledTimes(2));
    await vi.waitFor(() => {
      expect(el.querySelector("canvas")).not.toBeNull();
    });
    expect(onError).not.toHaveBeenCalled(); // 取消不是错误
    inst.destroy();
  });

  it("does not leak pendingRelease across a failed render (no misfire release on retry)", async () => {
    // 在途渲染期间收到释放请求（pendingRelease=true），随后渲染以非取消错误
    // 失败。若失败分支不清 pendingRelease，下一次成功渲染会在完成瞬间被
    // 误判为"过期页"立即释放，留下空白。
    const page = fakePage(1);
    const d = deferred<void>();
    let calls = 0;
    const render = vi.fn().mockImplementation(() => {
      calls++;
      return calls === 1
        ? { promise: d.promise, cancel: vi.fn() } // cancel 无操作：由 d 决定结局
        : { promise: Promise.resolve(), cancel: vi.fn() };
    });
    (page as unknown as { render: unknown }).render = render;
    const { doc } = fakeDoc([page]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onError,
    });
    await waitForPages(container, 1);
    const el = pageEls(container)[0];

    ioByMargin("100%")[0].fire(el, true);
    await vi.waitFor(() => expect(render).toHaveBeenCalledTimes(1));
    ioByMargin("300%")[0].fire(el, false); // pendingRelease = true
    d.reject(new Error("boom")); // 非取消错误
    await vi.waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
    expect(el.querySelector("canvas")).toBeNull();

    ioByMargin("100%")[0].fire(el, true); // 重入 → 重试
    await vi.waitFor(() => expect(render).toHaveBeenCalledTimes(2));
    // 成功渲染的页必须保留 canvas，而不是被残留的 pendingRelease 立即释放
    await vi.waitFor(() => {
      expect(el.querySelector("canvas")).not.toBeNull();
    });
    inst.destroy();
  });

  it("drops a queued page that leaves the render band before its turn", async () => {
    const pages = [fakePage(1, { deferred: true }), fakePage(2)];
    const { doc } = fakeDoc(pages);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, { source: magicSource() });
    await waitForPages(container, 2);
    const els = pageEls(container);

    // 页 1 占住渲染位（deferred 不放行），页 2 入队后立即滚出渲染带
    ioByMargin("100%")[0].fire(els[0], true);
    await vi.waitFor(() => expect(pages[0].render).toHaveBeenCalled());
    ioByMargin("100%")[0].fire(els[1], true);
    ioByMargin("100%")[0].fire(els[1], false); // wanted=false
    await new Promise((r) => setTimeout(r, 20));
    expect(pages[1].render).not.toHaveBeenCalled();
    inst.destroy();
  });

  /* --------------------------- error normalization --------------------------- */

  it("reports UNSUPPORTED before touching the engine (no getDocument call)", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    const notPdf = new Uint8Array([1, 2, 3, 4]);
    const inst = createPdfPreview(container, { source: notPdf, onError });
    await vi.waitFor(() => expect(onError).toHaveBeenCalled());
    const err = onError.mock.calls[0][0] as PdfPreviewError;
    expect(err.code).toBe("UNSUPPORTED");
    expect(getDocument).not.toHaveBeenCalled();
    expect(container.children).toHaveLength(0); // 失败即无副作用
    inst.destroy();
  });

  it.each([
    [1, /password-protected/i],
    [2, /rejected/i],
  ])(
    "normalizes PasswordException code %i to PASSWORD_PROTECTED",
    async (code, messageRe) => {
      vi.stubGlobal("IntersectionObserver", FakeIO);
      const onError = vi.fn();
      getDocumentMock.mockImplementation((() => ({
        promise: Promise.reject(
          new FakePasswordException("No password given", code),
        ),
        destroy: vi.fn(async () => {}),
      })) as never);
      const inst = createPdfPreview(container, {
        source: magicSource(),
        onError,
      });
      await vi.waitFor(() => expect(onError).toHaveBeenCalled());
      const err = onError.mock.calls[0][0] as PdfPreviewError;
      expect(err.code).toBe("PASSWORD_PROTECTED");
      expect(messageRe.test(err.message)).toBe(true);
      inst.destroy();
    },
  );

  it("normalizes InvalidPDFException to CORRUPT with cause preserved", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    const cause = new Error("Invalid PDF structure");
    cause.name = "InvalidPDFException";
    getDocumentMock.mockImplementation((() => ({
      promise: Promise.reject(cause),
      destroy: vi.fn(async () => {}),
    })) as never);
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onError,
    });
    await vi.waitFor(() => expect(onError).toHaveBeenCalled());
    const err = onError.mock.calls[0][0] as PdfPreviewError;
    expect(err.code).toBe("CORRUPT");
    expect(err.cause).toBe(cause);
    inst.destroy();
  });

  it("does not fire onRendered from an in-flight render after a document-level failure", async () => {
    // 首页渲染在途时其余页 getPage 失败 → teardown（onError 已上报、容器已清）。
    // 首页任务随后完成时不得再走成功路径触发 onRendered（失败后就该静默）。
    const first = fakePage(1);
    const d = deferred<void>();
    (first as unknown as { render: unknown }).render = vi.fn(() => ({
      promise: d.promise,
      cancel: vi.fn(),
    }));
    const page2 = deferred<FakePage>();
    const doc = {
      numPages: 2,
      getPage: (n: number) =>
        n === 1 ? Promise.resolve(first) : page2.promise,
    };
    getDocumentMock.mockImplementation((() => ({
      promise: Promise.resolve(doc),
      destroy: vi.fn(async () => {}),
    })) as never);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onRendered = vi.fn();
    const onError = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onRendered,
      onError,
    });
    await waitForPages(container, 1);
    ioByMargin("100%")[0].fire(pageEls(container)[0], true);
    await vi.waitFor(() => expect(first.render).toHaveBeenCalled());

    page2.reject(new Error("page 2 unavailable")); // 文档级失败
    await vi.waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
    d.resolve(); // 首页迟到完成
    await new Promise((r) => setTimeout(r, 20));
    expect(onRendered).not.toHaveBeenCalled();
    // 失败终态只报一次：迟到渲染不得再制造第二条错误
    expect(onError).toHaveBeenCalledTimes(1);
    inst.destroy();
  });

  it("reports CORRUPT and stays alive when a page render fails unexpectedly", async () => {
    const pages = [fakePage(1, { failRender: new Error("boom") })];
    const { doc } = fakeDoc(pages);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onError,
    });
    await waitForPages(container, 1);
    ioByMargin("100%")[0].fire(pageEls(container)[0], true);
    await vi.waitFor(() => expect(onError).toHaveBeenCalled());
    const err = onError.mock.calls[0][0] as PdfPreviewError;
    expect(err.code).toBe("CORRUPT");
    expect(String(err.message)).toContain("boom");
    // 状态机回 empty：可重试重渲
    ioByMargin("100%")[0].fire(pageEls(container)[0], true);
    await vi.waitFor(() => expect(pages[0].render).toHaveBeenCalledTimes(2));
    inst.destroy();
  });

  it("throws synchronously on an invalid numeric scale (programming error)", () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    expect(() =>
      createPdfPreview(container, { source: magicSource(), scale: -1 }),
    ).toThrow(/Invalid scale/);
    expect(getDocument).not.toHaveBeenCalled();
  });

  it("forwards whitelist renderOptions to getDocument and pins taken-over params", async () => {
    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, {
      source: magicSource(),
      renderOptions: { disableFontFace: true, useSystemFonts: false },
    });
    await waitForPages(container, 1);
    const params = getDocumentMock.mock.calls[0][0] as Record<string, unknown>;
    expect(params.disableFontFace).toBe(true);
    expect(params.useSystemFonts).toBe(false);
    expect(params.cMapPacked).toBe(true); // 接管项：bcmap 二进制格式
    expect(params.enableXfa).toBe(false); // 接管项：XFA 首版不支持
    expect(typeof params.cMapUrl).toBe("string");
    expect(typeof params.standardFontDataUrl).toBe("string");
    inst.destroy();
  });

  it("explicit workerSrc/cMapUrl/standardFontDataUrl win over derivation", async () => {
    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, {
      source: magicSource(),
      workerSrc: "/w.js",
      cMapUrl: "/cm/",
      standardFontDataUrl: "/sf/",
      wasmUrl: "/wasm/",
      iccUrl: "/icc/",
    });
    await waitForPages(container, 1);
    const params = getDocumentMock.mock.calls[0][0] as Record<string, unknown>;
    expect(params.cMapUrl).toBe("/cm/");
    expect(params.standardFontDataUrl).toBe("/sf/");
    expect(params.wasmUrl).toBe("/wasm/");
    expect(params.iccUrl).toBe("/icc/");
    if (typeof Worker !== "undefined") {
      expect(GlobalWorkerOptionsMock.workerSrc).toBe("/w.js");
    }
    inst.destroy();
  });

  it("injected stylesheet carries styleNonce and ships the text-layer variable chain", async () => {
    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, {
      source: magicSource(),
      styleNonce: "n-123",
    });
    await waitForPages(container, 1);
    const style = document.getElementById(
      "marcusok-pdf-preview-styles",
    ) as HTMLStyleElement | null;
    expect(style).not.toBeNull();
    expect(style!.getAttribute("nonce")).toBe("n-123");
    expect(style!.textContent).toContain("--total-scale-factor");
    expect(style!.textContent).toContain(".textLayer");
    inst.destroy();
  });

  /* ---------------------- source re-use (transfer safety) ---------------------- */

  // 上游 GetDocRequest 以 [data.buffer] 为 transfer list（pdf.mjs 源码实证）：
  // 引擎拿走字节后调用方持有的 source 立即 detach。toBytes 必须复制字节型
  // source，否则同一 source 二次 create（"重 create 即可"路径，play demo
  // 的切缩放/重建即此形态）会以 "File is empty." 假报错失败。
  it.each([
    ["Uint8Array", (b: Uint8Array) => b],
    ["ArrayBuffer", (b: Uint8Array) => b.buffer as ArrayBuffer],
  ])(
    "re-creating with the same %s source works (bytes are copied, not transferred)",
    async (_label, wrap) => {
      const pages = [fakePage(1)];
      const { doc } = fakeDoc(pages);
      // 模拟上游 transfer 语义：真把 params.data 的 buffer detach 掉
      getDocumentMock.mockImplementation((params: { data: Uint8Array }) => {
        structuredClone(params.data, { transfer: [params.data.buffer] });
        return {
          promise: Promise.resolve(doc),
          destroy: vi.fn(async () => {}),
        };
      });
      vi.stubGlobal("IntersectionObserver", FakeIO);
      const onError = vi.fn();
      const source = wrap(magicSource());

      const first = createPdfPreview(container, { source, onError });
      await waitForPages(container, 1);
      first.destroy();

      // 第二次 create 用同一 source：sniff 必须仍能读到完整魔数
      const second = createPdfPreview(container, { source, onError });
      await vi.waitFor(() => {
        expect(container.querySelectorAll(".mop-pdf-page")).toHaveLength(1);
      });
      expect(onError).not.toHaveBeenCalled();
      second.destroy();
    },
  );

  /* ------------------------- scroll-root selection ------------------------- */

  it("prefers a height-constrained scroll ancestor over an unconstrained overflow wrapper", async () => {
    // 复现 play 的 ant-card-body 形态：mid 是 overflow:auto 但高度随内容
    // 增长；旧实现会选 mid 作 IO root，root 矩形覆盖整份文档，所有页都
    // 判定为 intersecting（200 页 demo 全量渲染）。必须跳过 mid 选 outer。
    const outer = document.createElement("div");
    outer.style.overflow = "auto";
    outer.style.height = "300px";
    const mid = document.createElement("div");
    mid.style.overflow = "auto";
    const host = document.createElement("div");
    mid.appendChild(host);
    outer.appendChild(mid);
    document.body.appendChild(outer);

    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(host, { source: magicSource() });
    await waitForPages(host, 1);

    expect(ioByMargin("100%")[0].options?.root).toBe(outer);
    expect(ioByMargin("300%")[0].options?.root).toBe(outer);
    inst.destroy();
  });

  it("falls back to the viewport when no ancestor has a constrained height", async () => {
    const outer = document.createElement("div");
    outer.style.overflow = "auto"; // 无 height/max-height、当前也不溢出
    const host = document.createElement("div");
    outer.appendChild(host);
    document.body.appendChild(outer);

    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(host, { source: magicSource() });
    await waitForPages(host, 1);

    expect(ioByMargin("100%")[0].options?.root).toBeNull();
    expect(ioByMargin("300%")[0].options?.root).toBeNull();
    inst.destroy();
  });

  it("detects an overflow-only scroll ancestor once the first placeholder exists", async () => {
    // 类定义固定高度 / flex 解析高度的面板（无内联 height、无 max-height）
    // 在 create 时无从检出——唯一线索是占位内容撑出的真实溢出。观察器
    // 惰性创建后此形态可被选中为 root（happy-dom 无布局，用 getter 桩
    // 模拟"占位前不溢出、占位后溢出"；旧实现在 create 时刻探测，恒漏检）。
    const outer = document.createElement("div");
    outer.style.overflow = "auto"; // 无 height/max-height：溢出是唯一线索
    const host = document.createElement("div");
    outer.appendChild(host);
    document.body.appendChild(outer);
    Object.defineProperty(outer, "clientHeight", { get: () => 300 });
    Object.defineProperty(outer, "scrollHeight", {
      get: () => (outer.querySelector(".mop-pdf-page") ? 1000 : 0),
    });

    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(host, { source: magicSource() });
    await waitForPages(host, 1);

    expect(ioByMargin("100%")[0].options?.root).toBe(outer);
    expect(ioByMargin("300%")[0].options?.root).toBe(outer);
    inst.destroy();
  });

  /* ---------------------- worker/asset failure mapping ---------------------- */

  it("maps worker setup failures to UNKNOWN instead of CORRUPT", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    getDocumentMock.mockImplementation((() => ({
      promise: Promise.reject(
        new Error(
          'Setting up fake worker failed: "Failed to fetch dynamically imported module: /bad-worker.mjs".',
        ),
      ),
      destroy: vi.fn(async () => {}),
    })) as never);

    const inst = createPdfPreview(container, {
      source: magicSource(),
      onError,
    });
    await vi.waitFor(() => expect(onError).toHaveBeenCalled());
    const err = onError.mock.calls[0][0] as PdfPreviewError;
    expect(err.code).toBe("UNKNOWN");
    expect(err.message).toMatch(/worker/i);
    inst.destroy();
  });

  /* --------------------- text layer / page resource cleanup --------------------- */

  it("cancels an in-flight text layer and cleans page resources on release", async () => {
    const pages = [fakePage(1)];
    const { doc } = fakeDoc(pages);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);

    const tlRender = deferred<void>();
    const tlCancel = vi.fn(() => {
      tlRender.reject(
        Object.assign(new Error("TextLayer task cancelled."), {
          name: "AbortException",
        }),
      );
    });
    TextLayerMock.mockImplementationOnce(function fakeTextLayer() {
      return { render: vi.fn(() => tlRender.promise), cancel: tlCancel };
    });

    const inst = createPdfPreview(container, { source: magicSource() });
    await waitForPages(container, 1);
    const el = pageEls(container)[0];

    ioByMargin("100%")[0].fire(el, true);
    await vi.waitFor(() => expect(TextLayerMock).toHaveBeenCalled());
    // canvas 已完成、文本层仍在抽取：离开保留带应取消文本流并清页资源
    ioByMargin("300%")[0].fire(el, false);

    await vi.waitFor(() => expect(tlCancel).toHaveBeenCalledTimes(1));
    await vi.waitFor(() => {
      expect(el.querySelector("canvas")).toBeNull();
      expect(el.querySelector(".textLayer")).toBeNull();
    });
    expect(pages[0].cleanup).toHaveBeenCalled();
    inst.destroy();
  });

  it("zeroes rendered canvases and clears the wrapper on destroy", async () => {
    const pages = [fakePage(1)];
    const { doc } = fakeDoc(pages);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, {
      source: magicSource(),
      scale: 1,
    });
    await waitForPages(container, 1);
    const el = pageEls(container)[0];
    ioByMargin("100%")[0].fire(el, true);
    await vi.waitFor(() => {
      expect(el.querySelector("canvas")).not.toBeNull();
    });
    const canvas = el.querySelector("canvas")!;
    expect(canvas.width).toBe(612);

    inst.destroy();
    expect(canvas.width).toBe(0);
    expect(canvas.height).toBe(0);
    expect(container.children).toHaveLength(0);
    expect(pages[0].cleanup).toHaveBeenCalled();
  });

  /* ------------------------------- sources ------------------------------- */

  it("accepts a Blob source", async () => {
    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    // 用最小 Blob 形态直测 toBytes 的 arrayBuffer() 分支
    const blob = {
      arrayBuffer: () => Promise.resolve(magicSource().buffer as ArrayBuffer),
    } as unknown as Blob;
    const inst = createPdfPreview(container, { source: blob });
    await waitForPages(container, 1);
    inst.destroy();
  });

  it("throws synchronously on an invalid scale string", () => {
    expect(() =>
      createPdfPreview(container, {
        source: magicSource(),
        scale: "fit" as never,
      }),
    ).toThrow(/Invalid scale/);
    expect(getDocument).not.toHaveBeenCalled();
  });
});

describe("用户回调异常隔离", () => {
  it("异步回调抛错被兜住（仅 warn，不产生 unhandled rejection）", async () => {
    // TS 允许把 async 函数赋给返回 void 的回调类型：其 rejected promise
    // 不会被同步 catch 捕到——guardCallback 必须显式 .catch
    const { doc } = fakeDoc([fakePage(1)]);
    setupDocument(doc);
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const inst = createPdfPreview(container, {
      source: magicSource(),
      // eslint-disable-next-line @typescript-eslint/no-misused-promises -- 有意模拟 async void 回调的 rejected promise
      onRendered: async () => {
        await Promise.resolve();
        throw new Error("async cb boom");
      },
    });
    await waitForPages(container, 1);
    ioByMargin("100%")[0].fire(pageEls(container)[0], true);
    await vi.waitFor(() => {
      expect(
        warn.mock.calls.some((c) => String(c[0]).includes("onRendered")),
      ).toBe(true);
    });
    inst.destroy();
    warn.mockRestore();
  });
});
