// @vitest-environment happy-dom
/**
 * destroy 时 TextLayer.cleanup 的微任务时机（2026-11 修复）：
 * 上游 pdf.mjs（6.4 实测）的语义是——实例构造时进入模块级
 * #pendingTextLayers；capability settle（含 cancel）后其自移除回调排在
 * **微任务**里；静态 cleanup() 见 pending 非空直接 return。teardown 在
 * 同一同步块里 cancel 全部在途文本层后同步调 cleanup，必然被短路，此后
 * 无人再调——测量 canvas 与 ascent 缓存常驻到页面卸载。
 *
 * 修复：cleanup 推迟到微任务（排在全部 cancel 的移除回调之后），回调内
 * 复查活跃实例计数（同帧新建的实例不被误清）。本文件的 TextLayer mock
 * 精确复刻上游三点语义（pending 集 / 微任务移除 / 非空短路）。
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

/** 上游静态语义的观测面：pending 集、cleanup 调用数、真正清掉的次数。 */
const textLayerState = vi.hoisted(() => ({
  pending: new Set<unknown>(),
  cleanupCalls: 0,
  cleaned: 0,
}));

const TextLayerMock = vi.hoisted(() =>
  vi.fn(function fakeTextLayer() {
    let settle!: () => void;
    const capability = new Promise<void>((r) => {
      settle = r;
    });
    const inst = {
      // 在途文本层：render 永不落定（destroy 场景的典型形态——大页文本
      // 流式提取中）
      render: vi.fn(async () => {
        await new Promise(() => {});
      }),
      // cancel 同步 settle capability → 自移除排在下一个微任务（上游同款）
      cancel: vi.fn(() => settle()),
    };
    // 上游语义：移除发生在 settle 后的微任务（lint 的 no-floating-promises
    // 要求显式 void——该 promise 永不 reject，无需 rejection 处理）
    void capability.then(() => {
      textLayerState.pending.delete(inst);
    });
    textLayerState.pending.add(inst);
    return inst;
  }),
);
// 静态 cleanup（生产代码以 TextLayer 为接收者调用）：pending 非空短路
(TextLayerMock as unknown as { cleanup: () => void }).cleanup = () => {
  textLayerState.cleanupCalls++;
  if (textLayerState.pending.size > 0) return;
  textLayerState.cleaned++;
};

vi.mock("pdfjs-dist", () => ({
  getDocument: getDocumentMock,
  GlobalWorkerOptions: { workerSrc: "" },
  TextLayer: TextLayerMock,
  InvalidPDFException: class extends Error {},
  PasswordException: class extends Error {},
  RenderingCancelledException: class extends Error {},
}));

function magicSource(): Uint8Array {
  const bytes = new Uint8Array(64);
  bytes.set([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x34], 0);
  return bytes;
}

function fakePage(num: number) {
  return {
    pageNumber: num,
    getViewport: ({ scale = 1 }: { scale?: number }) => ({
      width: 612 * scale,
      height: 792 * scale,
      scale,
      userUnit: 1,
    }),
    render: vi.fn(() => ({ promise: Promise.resolve(), cancel: vi.fn() })),
    cleanup: vi.fn(() => true),
    streamTextContent: () => ({ fake: "stream" }),
  };
}

function setupOnePager() {
  getDocumentMock.mockImplementation(() => ({
    promise: Promise.resolve({
      numPages: 1,
      getPage: () => Promise.resolve(fakePage(1)),
    }),
    destroy: vi.fn(async () => {}),
  }));
}

/** 走完微任务队列（cancel 的移除回调 + 排队的 cleanup 都落定）。 */
async function drainMicrotasks(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

/** 取渲染观察器（margin "100%"，与 render.test.ts 同口径）。 */
function renderIo(): FakeIO {
  return FakeIO.instances.filter((io) =>
    String(io.options?.rootMargin ?? "").startsWith("100%"),
  )[0]!;
}

function pageEls(container: HTMLElement): HTMLElement[] {
  return Array.from(container.querySelectorAll<HTMLElement>(".mop-pdf-page"));
}

/** 触发首页渲染（占位存在后手动派发渲染观察器回调）。 */
async function fireFirstPage(container: HTMLElement): Promise<void> {
  await vi.waitFor(() => {
    expect(pageEls(container)).toHaveLength(1);
  });
  renderIo().fire(pageEls(container)[0]!, true);
}

describe("destroy 时 TextLayer.cleanup 时机", () => {
  beforeEach(() => {
    document.body.textContent = "";
    FakeIO.instances = [];
    getDocumentMock.mockReset();
    TextLayerMock.mockClear();
    textLayerState.pending.clear();
    textLayerState.cleanupCalls = 0;
    textLayerState.cleaned = 0;
    vi.stubGlobal("IntersectionObserver", FakeIO);
  });

  it("在途文本层时 destroy：cleanup 延后一个微任务后真正执行", async () => {
    setupOnePager();
    const container = document.createElement("div");
    document.body.appendChild(container);
    Object.defineProperty(container, "clientWidth", {
      value: 600,
      configurable: true,
    });
    Object.defineProperty(container, "clientHeight", {
      value: 800,
      configurable: true,
    });

    const preview = createPdfPreview(container, { source: magicSource() });
    // 首页渲染推进到文本层在途（canvas 渲染即成，tl.render 永不落定）
    await fireFirstPage(container);
    await vi.waitFor(() => {
      expect(textLayerState.pending.size).toBe(1);
    });

    preview.destroy();
    // 同步块内不得调 cleanup（此前的同步调用必被 pending 短路，缓存常驻）
    expect(textLayerState.cleanupCalls).toBe(0);
    await drainMicrotasks();
    // 微任务后：移除回调先出队（pending 清空）、cleanup 随后真正清缓存
    expect(textLayerState.pending.size).toBe(0);
    expect(textLayerState.cleanupCalls).toBe(1);
    expect(textLayerState.cleaned).toBe(1);
  });

  it("destroy 后同帧新建实例：延迟的 cleanup 复查计数不清新实例缓存", async () => {
    setupOnePager();
    const containerA = document.createElement("div");
    document.body.appendChild(containerA);
    Object.defineProperty(containerA, "clientWidth", {
      value: 600,
      configurable: true,
    });
    Object.defineProperty(containerA, "clientHeight", {
      value: 800,
      configurable: true,
    });
    const a = createPdfPreview(containerA, { source: magicSource() });
    await fireFirstPage(containerA);
    await vi.waitFor(() => {
      expect(textLayerState.pending.size).toBe(1);
    });
    a.destroy();

    // 同帧（微任务跑之前）新建实例 B：延迟的 cleanup 被活跃计数复查拦下，
    // 根本不调用（避免清掉存活实例的缓存）
    const containerB = document.createElement("div");
    document.body.appendChild(containerB);
    Object.defineProperty(containerB, "clientWidth", {
      value: 600,
      configurable: true,
    });
    Object.defineProperty(containerB, "clientHeight", {
      value: 800,
      configurable: true,
    });
    const b = createPdfPreview(containerB, { source: magicSource() });
    await drainMicrotasks();
    expect(textLayerState.cleanupCalls).toBe(0); // 复查拦下，未调用

    // B 销毁后（pending 已空、计数归零）cleanup 才真正执行
    b.destroy();
    await drainMicrotasks();
    expect(textLayerState.cleanupCalls).toBe(1);
    expect(textLayerState.cleaned).toBe(1);
  });

  it("无在途文本层的 destroy：cleanup 同样经微任务执行（行为不劣化）", async () => {
    setupOnePager();
    const container = document.createElement("div");
    document.body.appendChild(container);
    Object.defineProperty(container, "clientWidth", {
      value: 600,
      configurable: true,
    });
    Object.defineProperty(container, "clientHeight", {
      value: 800,
      configurable: true,
    });
    const preview = createPdfPreview(container, {
      source: magicSource(),
      textLayer: false,
    });
    await vi.waitFor(() => {
      expect(container.querySelectorAll(".mop-pdf-page")).toHaveLength(1);
    });
    preview.destroy();
    expect(textLayerState.cleanupCalls).toBe(0);
    await drainMicrotasks();
    expect(textLayerState.cleanupCalls).toBe(1);
    expect(textLayerState.cleaned).toBe(1);
  });
});
