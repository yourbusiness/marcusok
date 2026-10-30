// @vitest-environment happy-dom
/**
 * 安全加固测试：超链接 scheme 白名单、altChunk sandbox、styleNonce 换节点
 * 重建、默认关闭 altChunk、选项接管与 destroy 静默。
 *
 * 为什么 mock docx-preview（计划第七轮 P0）：happy-dom 的 DOMParser 对
 * 带前缀属性（r:id）保留完整 localName，上游 renderHyperlink 读不到
 * `r:id` → href 恒为空——真实 r:id fixture 在 happy-dom 下**没有 sanitizer
 * 也会通过**（假阳性）。因此安全断言不依赖上游真实解析，由 mock 直接向
 * 容器注入恶意产物来验证本包后处理；真实 r:id 路径在浏览器 demo 复核。
 * sniff 仍走真实路径（不依赖 docx-preview），fixture 须为合法四文件 docx。
 */
import { describe, expect, it, vi, beforeEach } from "vitest";

// vi.mock 工厂被提升到文件顶部，共享 mock 须经 vi.hoisted（否则引用到
// 未初始化的变量）
const { renderAsyncMock } = vi.hoisted(() => ({ renderAsyncMock: vi.fn() }));

vi.mock("docx-preview", () => ({
  renderAsync: renderAsyncMock,
}));

import { createWordPreview } from "../index";
import type { Options } from "docx-preview";
import { simpleDocx } from "./fixtures";

/** mock 收到的上游选项（第 4 参）。 */
function lastUpstreamOptions(): Partial<Options> {
  return renderAsyncMock.mock.calls.at(-1)![3] as Partial<Options>;
}

/** 状态轮询等待（替代固定 sleep；与 render.test.ts 同款，见其注释）。 */
async function waitFor(cond: () => boolean, timeoutMs = 2000): Promise<void> {
  const t0 = Date.now();
  while (!cond()) {
    if (Date.now() - t0 > timeoutMs)
      throw new Error("waitFor: condition not met in time");
    await new Promise((res) => setTimeout(res, 10));
  }
}

function makeContainer(): HTMLElement {
  document.body.textContent = "";
  const c = document.createElement("div");
  document.body.appendChild(c);
  return c;
}

/** 让 mock renderAsync 把给定 HTML 写进 bodyContainer（模拟上游渲染产物）。 */
function mockRender(html: string): void {
  renderAsyncMock.mockImplementation((_data, bodyContainer: HTMLElement) => {
    bodyContainer.innerHTML = html;
  });
}

beforeEach(() => {
  renderAsyncMock.mockReset();
});

describe("超链接 scheme 白名单", () => {
  it("javascript: href 被摘除、链接文本保留", async () => {
    mockRender(`<a href="javascript:alert(1)">click me</a>`);
    const container = makeContainer();
    createWordPreview(container, { source: simpleDocx("x") });
    await waitFor(() => container.querySelector("a") !== null);

    const a = container.querySelector("a")!;
    expect(a.getAttribute("href")).toBeNull();
    expect(a.textContent).toBe("click me");
  });

  it("data: 与 vbscript: 拒绝；https:/http:/mailto:/tel:/#锚点/相对路径 放行", async () => {
    mockRender(
      `<a href="data:text/html,<script>alert(1)</script>" id="a1">d</a>` +
        `<a href="vbscript:msgbox(1)" id="a2">v</a>` +
        `<a href="https://example.com/x" id="a3">ok1</a>` +
        `<a href="http://example.com/" id="a4">ok2</a>` +
        `<a href="mailto:a@b.c" id="a5">ok3</a>` +
        `<a href="tel:+1234" id="a6">ok4</a>` +
        `<a href="#section1" id="a7">ok5</a>` +
        `<a href="docs/page.docx" id="a8">ok6</a>`,
    );
    const container = makeContainer();
    createWordPreview(container, { source: simpleDocx("x") });
    await waitFor(() => container.querySelector("#a8") !== null);

    expect(container.querySelector("#a1")!.getAttribute("href")).toBeNull();
    expect(container.querySelector("#a2")!.getAttribute("href")).toBeNull();
    for (const id of ["a3", "a4", "a5", "a6", "a7", "a8"]) {
      expect(
        container.querySelector(`#${id}`)!.getAttribute("href"),
      ).toBeTruthy();
    }
  });

  it("大小写与前导空白绕过被拦截（URL 解析归一 + 控制字符剔除）", async () => {
    mockRender(
      `<a href="JaVaScRiPt:alert(1)" id="c1">x</a>` +
        `<a href="  javascript:alert(1)" id="c2">x</a>` +
        `<a href="&#01;javascript:alert(1)" id="c3">x</a>` +
        // 对照：前导空白的 https 必须仍放行（不能误杀）
        `<a href="  https://example.com/" id="c4">x</a>`,
    );
    const container = makeContainer();
    createWordPreview(container, { source: simpleDocx("x") });
    await waitFor(() => container.querySelector("#c4") !== null);

    expect(container.querySelector("#c1")!.getAttribute("href")).toBeNull();
    expect(container.querySelector("#c2")!.getAttribute("href")).toBeNull();
    expect(container.querySelector("#c3")!.getAttribute("href")).toBeNull();
    expect(container.querySelector("#c4")!.getAttribute("href")).toBeTruthy();
  });
});

describe("altChunk 加固", () => {
  it("默认 renderAltChunks: false（上游默认是 true）", async () => {
    mockRender("<p>x</p>");
    const container = makeContainer();
    createWordPreview(container, { source: simpleDocx("x") });
    await waitFor(() => container.querySelector("p") !== null);

    expect(lastUpstreamOptions().renderAltChunks).toBe(false);
  });

  it("allowAltChunks: true → 透传开启，渲染后 iframe 补 sandbox", async () => {
    mockRender(`<iframe srcdoc="<script>alert(1)</script>"></iframe>`);
    const container = makeContainer();
    createWordPreview(container, {
      source: simpleDocx("x"),
      allowAltChunks: true,
    });
    await waitFor(() => container.querySelector("iframe") !== null);

    expect(lastUpstreamOptions().renderAltChunks).toBe(true);
    expect(container.querySelector("iframe")!.getAttribute("sandbox")).toBe("");
  });

  it("默认路径不补 sandbox（后处理仅在开启 altChunk 时动 iframe）", async () => {
    mockRender(`<iframe srcdoc="x"></iframe>`);
    const container = makeContainer();
    // 默认 allowAltChunks=false：上游被关闭、本不产 iframe；但若其他路径
    // 出现 iframe（如未来版本行为变化），sandbox 仅在开启 altChunk 时补
    // ——验证默认路径既不透传开启、也不误加属性
    createWordPreview(container, { source: simpleDocx("x") });
    await waitFor(() => container.querySelector("iframe") !== null);
    expect(lastUpstreamOptions().renderAltChunks).toBe(false);
    expect(
      container.querySelector("iframe")!.getAttribute("sandbox"),
    ).toBeNull();
  });
});

describe("styleNonce（换节点重建）", () => {
  it("有 nonce：上游注入的 style 被重建为带 nonce 的新节点并挂回渲染容器", async () => {
    renderAsyncMock.mockImplementation(
      (_data, bodyContainer: HTMLElement, styleContainer?: HTMLElement) => {
        // 模拟上游：STYLE 节点路由到 styleContainer（游离容器）
        const style = document.createElement("style");
        style.textContent = ".docx-1 p { margin: 0 }";
        if (styleContainer) styleContainer.appendChild(style);
        bodyContainer.innerHTML = "<p>x</p>";
      },
    );
    const container = makeContainer();
    createWordPreview(container, {
      source: simpleDocx("x"),
      styleNonce: "abc123",
    });
    await waitFor(() => container.querySelector("style") !== null);

    // 渲染容器内最终有带 nonce 的重建节点，CSS 文本一致
    const style = container.querySelector("style")!;
    expect(style).toBeTruthy();
    expect(style.getAttribute("nonce")).toBe("abc123");
    expect(style.textContent).toBe(".docx-1 p { margin: 0 }");
  });

  it("无 nonce：不传第三参（上游并入 bodyContainer，不多走游离重建）", async () => {
    let styleContainerSeen: unknown = "unset";
    renderAsyncMock.mockImplementation(
      (_data, bodyContainer: HTMLElement, styleContainer?: HTMLElement) => {
        styleContainerSeen = styleContainer;
        const style = document.createElement("style");
        style.textContent = "p{}";
        bodyContainer.appendChild(style);
      },
    );
    const container = makeContainer();
    createWordPreview(container, { source: simpleDocx("x") });
    await waitFor(() => container.querySelector("style") !== null);

    expect(styleContainerSeen).toBeUndefined();
    expect(container.querySelector("style")!.getAttribute("nonce")).toBeNull();
  });
});

describe("默认加固选项接管", () => {
  it("上游收到：唯一 className、useBase64URL: true、ignoreLastRenderedPageBreak: false", async () => {
    mockRender("<p>x</p>");
    const container = makeContainer();
    createWordPreview(container, { source: simpleDocx("x") });
    await waitFor(() => container.querySelector("p") !== null);

    const opts = lastUpstreamOptions();
    expect(opts.className).toMatch(/^docx-\d+$/);
    expect(opts.useBase64URL).toBe(true);
    expect(opts.ignoreLastRenderedPageBreak).toBe(false);
  });

  it("白名单透传：renderOptions 的 13 项直传上游；ignoreLastRenderedPageBreak 可改回 true", async () => {
    mockRender("<p>x</p>");
    const container = makeContainer();
    createWordPreview(container, {
      source: simpleDocx("x"),
      renderOptions: {
        breakPages: false,
        ignoreWidth: true,
        renderComments: true,
        ignoreLastRenderedPageBreak: true,
      },
    });
    await waitFor(() => container.querySelector("p") !== null);

    const opts = lastUpstreamOptions();
    expect(opts.breakPages).toBe(false);
    expect(opts.ignoreWidth).toBe(true);
    expect(opts.renderComments).toBe(true);
    expect(opts.ignoreLastRenderedPageBreak).toBe(true); // 用户显式改回
    // 接管项仍不被透传覆盖
    expect(opts.className).toMatch(/^docx-\d+$/);
    expect(opts.useBase64URL).toBe(true);
  });
});

describe("destroy 静默与错误归一", () => {
  it("destroy 后 renderAsync 才 resolve：结果丢弃、无任何回调、容器干净", async () => {
    let release!: () => void;
    renderAsyncMock.mockImplementation(
      (_data, bodyContainer: HTMLElement) =>
        new Promise<void>((resolve) => {
          release = () => {
            bodyContainer.innerHTML = "<p>late</p><style>p{}</style>";
            resolve();
          };
        }),
    );
    const container = makeContainer();
    const onRendered = vi.fn();
    const onError = vi.fn();
    const inst = createWordPreview(container, {
      source: simpleDocx("x"),
      onRendered,
      onError,
    });
    // boot 链（toBytes → sniff）先推进到 mock 挂起处：renderAsync 已在途
    // （destroy 提前则 boot 在 sniff 前就 return，mock 不会被调用）。
    // 轮询 mock 被调用（比固定 20ms 窗口稳）
    await waitFor(() => renderAsyncMock.mock.calls.length > 0);
    inst.destroy();
    release(); // 迟到的渲染完成（写入已摘除的私有 wrapper）
    // 静默路径无正向信号（回调按设计静默）：给足窗口（mock 已 resolve，
    // 剩余的是微任务边界，150ms 远超所需）
    await new Promise((res) => setTimeout(res, 150));

    expect(container.children.length).toBe(0);
    expect(onRendered).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
  });

  it("mock 抛错：归一 CORRUPT 并保留 cause", async () => {
    // 返回 rejected promise（同步抛出会被 mock 捕获、到不了 boot 的 catch）
    renderAsyncMock.mockImplementation(() =>
      Promise.reject(new Error("boom inside docx-preview")),
    );
    const container = makeContainer();
    const onError = vi.fn();
    createWordPreview(container, { source: simpleDocx("x"), onError });
    await waitFor(() => onError.mock.calls.length > 0);

    expect(onError).toHaveBeenCalledTimes(1);
    const err = onError.mock.calls[0][0];
    expect(err.code).toBe("CORRUPT");
    expect(err.message).toContain("boom inside docx-preview");
    expect(err.cause).toBeInstanceOf(Error);
    // catch 终止路径同样不残留私有 wrapper（失败即无副作用）
    expect(container.children.length).toBe(0);
  });

  it("渲染失败路径同样清理 CSS.highlights 注册（renderComments 半途失败不留残留）", async () => {
    const highlights = new Map<string, unknown>();
    vi.stubGlobal("CSS", { highlights });
    renderAsyncMock.mockImplementation(() =>
      Promise.reject(new Error("boom after highlight registered")),
    );
    const container = makeContainer();
    const onError = vi.fn();
    const inst = createWordPreview(container, {
      source: simpleDocx("x"),
      onError,
    });
    // className 在 create 时同步生成（wrapper 立即挂载），失败移除前先取
    const className = (container.firstElementChild as HTMLElement).className;
    highlights.set(`${className}-comments`, { fake: "Highlight" });
    await waitFor(() => onError.mock.calls.length > 0);
    expect(highlights.has(`${className}-comments`)).toBe(false);
    inst.destroy();
    vi.unstubAllGlobals();
  });

  it("输入源读取失败（Blob.arrayBuffer 抛错）：归 UNKNOWN 而非 CORRUPT", async () => {
    // toBytes 走真实路径（不经 docx-preview mock）；构造 arrayBuffer 拒绝的
    // 假 Blob 模拟"源读取失败"——与文档内容无关，按 api 文档口径归 UNKNOWN
    const badBlob = {
      arrayBuffer() {
        return Promise.reject(new Error("read source failed"));
      },
    } as unknown as Blob;
    const container = makeContainer();
    const onError = vi.fn();
    createWordPreview(container, { source: badBlob, onError });
    await waitFor(() => onError.mock.calls.length > 0);

    expect(onError).toHaveBeenCalledTimes(1);
    const err = onError.mock.calls[0][0];
    expect(err.code).toBe("UNKNOWN");
    expect(err.message).toContain("read source failed");
    expect(err.cause).toBeInstanceOf(Error);
    // 失败即无副作用：不残留私有 wrapper
    expect(container.children.length).toBe(0);
  });
});
