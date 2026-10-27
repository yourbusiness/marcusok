// @vitest-environment happy-dom
/**
 * destroy 全链路取消测试（mock 引擎，同 render.test 的环境边界）：
 * destroy 立即生效（容器清空、IO 断开、loadingTask.destroy 调用）、
 * 竞态静默（在途 getDocument/render/Page 完成后无回调无 DOM 写入）、幂等。
 */
import { describe, it, expect, vi, beforeEach } from "vitest";
import { createPdfPreview } from "../render/render";

class FakeIO {
  static instances: FakeIO[] = [];
  callback: IntersectionObserverCallback;
  options: IntersectionObserverInit | undefined;
  disconnected = false;
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
  disconnect() {
    this.disconnected = true;
  }
  fire(el: Element, isIntersecting: boolean) {
    this.callback(
      [{ target: el, isIntersecting } as IntersectionObserverEntry],
      this as unknown as IntersectionObserver,
    );
  }
}

const getDocumentMock = vi.hoisted(() => vi.fn());
vi.mock("pdfjs-dist", () => ({
  getDocument: getDocumentMock,
  GlobalWorkerOptions: { workerSrc: "" },
  TextLayer: vi.fn(function () {
    return { render: vi.fn(async () => {}), cancel: vi.fn() };
  }),
  InvalidPDFException: class extends Error {},
  PasswordException: class extends Error {},
  RenderingCancelledException: class extends Error {},
}));

function deferred<T>() {
  let resolve!: (v: T) => void;
  let reject!: (e: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

function magicSource(): Uint8Array {
  const bytes = new Uint8Array(64);
  bytes.set([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x34], 0);
  return bytes;
}

function fakePage(num: number, opts: { deferredRender?: boolean } = {}) {
  const d = deferred<void>();
  const tasks: { cancel: ReturnType<typeof vi.fn> }[] = [];
  return {
    pageNumber: num,
    getViewport: ({ scale = 1 }: { scale?: number }) => ({
      width: 612 * scale,
      height: 792 * scale,
      scale,
    }),
    render: vi.fn(() => {
      const task = {
        promise: opts.deferredRender ? d.promise : Promise.resolve(),
        cancel: vi.fn(),
      };
      tasks.push(task);
      return task;
    }),
    tasks,
    release: d,
    cleanup: vi.fn(() => true),
    streamTextContent: () => ({}),
  };
}

let container: HTMLElement;

beforeEach(() => {
  document.body.textContent = "";
  container = document.createElement("div");
  document.body.appendChild(container);
  document.getElementById("marcusok-pdf-preview-styles")?.remove();
  FakeIO.instances = [];
  getDocumentMock.mockReset();
});

describe("destroy", () => {
  it("destroys immediately: container emptied, observers disconnected, loadingTask destroyed", async () => {
    const destroy = vi.fn(async () => {});
    getDocumentMock.mockImplementation(() => ({
      promise: Promise.resolve({
        numPages: 1,
        getPage: () => Promise.resolve(fakePage(1)),
      }),
      destroy,
    }));
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onRendered = vi.fn();
    const onError = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onRendered,
      onError,
    });
    await vi.waitFor(() =>
      expect(container.querySelectorAll(".mop-pdf-page")).toHaveLength(1),
    );
    inst.destroy();
    expect(container.children).toHaveLength(0);
    expect(FakeIO.instances.every((io) => io.disconnected)).toBe(true);
    await vi.waitFor(() => expect(destroy).toHaveBeenCalledTimes(1));
    expect(onRendered).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
  });

  it("is idempotent", async () => {
    const destroy = vi.fn(async () => {});
    getDocumentMock.mockImplementation(() => ({
      promise: Promise.resolve({
        numPages: 1,
        getPage: () => Promise.resolve(fakePage(1)),
      }),
      destroy,
    }));
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const inst = createPdfPreview(container, { source: magicSource() });
    await vi.waitFor(() =>
      expect(container.querySelectorAll(".mop-pdf-page")).toHaveLength(1),
    );
    inst.destroy();
    inst.destroy();
    expect(destroy).toHaveBeenCalledTimes(1);
  });

  it("stays silent when getDocument settles after destroy", async () => {
    const d = deferred<{
      numPages: number;
      getPage: (n: number) => Promise<unknown>;
    }>();
    const destroy = vi.fn(async () => {});
    getDocumentMock.mockImplementation(() => ({
      promise: d.promise,
      destroy,
    }));
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onRendered = vi.fn();
    const onError = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onRendered,
      onError,
    });
    inst.destroy();
    // destroy 之后 getDocument 才 resolve：迟到结果全部丢弃
    d.resolve({ numPages: 1, getPage: () => Promise.resolve(fakePage(1)) });
    await new Promise((r) => setTimeout(r, 20));
    expect(container.children).toHaveLength(0);
    expect(onRendered).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
  });

  it("cancels in-flight render tasks and swallows the rejection", async () => {
    const page = fakePage(1, { deferredRender: true });
    getDocumentMock.mockImplementation(() => ({
      promise: Promise.resolve({
        numPages: 1,
        getPage: () => Promise.resolve(page),
      }),
      destroy: vi.fn(async () => {}),
    }));
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onError,
    });
    await vi.waitFor(() =>
      expect(container.querySelectorAll(".mop-pdf-page")).toHaveLength(1),
    );
    const el = container.querySelector(".mop-pdf-page") as HTMLDivElement;
    FakeIO.instances
      .find((io) => String(io.options?.rootMargin).startsWith("100%"))!
      .fire(el, true);
    await vi.waitFor(() => expect(page.render).toHaveBeenCalled());
    inst.destroy();
    // destroy 显式取消在途 RenderTask
    expect(page.tasks[0].cancel).toHaveBeenCalled();
    // destroy 触发取消语义：上游以 RenderingCancelledException 拒绝
    page.release.reject(
      Object.assign(new Error("Rendering cancelled"), {
        name: "RenderingCancelledException",
      }),
    );
    await new Promise((r) => setTimeout(r, 20));
    expect(onError).not.toHaveBeenCalled(); // 取消不是错误
    expect(container.children).toHaveLength(0);
  });

  it("stays silent when page enumeration settles after destroy", async () => {
    const d = deferred<void>();
    getDocumentMock.mockImplementation(() => ({
      promise: Promise.resolve({
        numPages: 3,
        getPage: (n: number) =>
          n === 1
            ? Promise.resolve(fakePage(1))
            : d.promise.then(() => fakePage(n)),
      }),
      destroy: vi.fn(async () => {}),
    }));
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onError,
    });
    await vi.waitFor(() =>
      expect(container.querySelectorAll(".mop-pdf-page")).toHaveLength(1),
    );
    inst.destroy();
    d.resolve();
    await new Promise((r) => setTimeout(r, 20));
    expect(container.children).toHaveLength(0);
    expect(onError).not.toHaveBeenCalled();
  });

  it("silences a source-read failure after destroy (no UNKNOWN callback)", async () => {
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    const pending = deferred<ArrayBuffer>();
    const inst = createPdfPreview(container, {
      // Blob 读取挂起，destroy 后才失败
      source: {
        arrayBuffer: () => pending.promise,
      } as unknown as Blob,
      onError,
    });
    inst.destroy();
    pending.reject(new Error("read failed"));
    await new Promise((r) => setTimeout(r, 20));
    expect(onError).not.toHaveBeenCalled();
  });

  it("cleans up worker and observers when boot fails (no destroy required)", async () => {
    const destroy = vi.fn(async () => {});
    const cause = new Error("Invalid PDF structure");
    cause.name = "InvalidPDFException";
    getDocumentMock.mockImplementation(() => ({
      promise: Promise.reject(cause),
      destroy,
    }));
    vi.stubGlobal("IntersectionObserver", FakeIO);
    const onError = vi.fn();
    const inst = createPdfPreview(container, {
      source: magicSource(),
      onError,
    });
    await vi.waitFor(() => expect(onError).toHaveBeenCalled());
    // 文档级失败即终止 worker 并断开观察器——不等消费方 destroy
    await vi.waitFor(() => expect(destroy).toHaveBeenCalledTimes(1));
    expect(FakeIO.instances.every((io) => io.disconnected)).toBe(true);
    expect(container.children).toHaveLength(0);
    // 其后 destroy() 仍安全且幂等：不二次终止 worker
    inst.destroy();
    inst.destroy();
    await new Promise((r) => setTimeout(r, 20));
    expect(destroy).toHaveBeenCalledTimes(1);
  });
});
