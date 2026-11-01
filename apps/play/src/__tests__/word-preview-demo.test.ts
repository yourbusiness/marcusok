// @vitest-environment happy-dom
/**
 * word-preview demo 回归：连续点击“合成大文档”必须重新挂载预览。
 *
 * 旧实现只更新 bytesRef/fileName，不 bump renderSeq；同名 fileName 会被
 * React bail-out 吞掉，effect 不重跑——失败后无法用该按钮重试。
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const { createWordPreviewMock } = vi.hoisted(() => ({
  createWordPreviewMock: vi.fn(),
}));

vi.mock("@marcusok/word-preview", () => ({
  createWordPreview: createWordPreviewMock,
}));

import WordPreviewDemo from "../demos/word-preview/word.demo";

async function waitFor(cond: () => boolean, timeoutMs = 3000): Promise<void> {
  const t0 = Date.now();
  while (!cond()) {
    if (Date.now() - t0 > timeoutMs) {
      throw new Error("waitFor: condition not met in time");
    }
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

describe("word-preview demo", () => {
  let root: Root | undefined;
  let host: HTMLElement;

  beforeEach(() => {
    createWordPreviewMock.mockReset();
    createWordPreviewMock.mockImplementation(() => ({ destroy: vi.fn() }));
    vi.stubGlobal("IS_REACT_ACT_ENVIRONMENT", true);
    vi.stubGlobal("matchMedia", (query: string) => ({
      matches: false,
      media: query,
      addListener: vi.fn(),
      removeListener: vi.fn(),
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
      dispatchEvent: vi.fn(),
    }));
    host = document.createElement("div");
    document.body.appendChild(host);
  });

  afterEach(async () => {
    if (root) {
      await act(async () => {
        root!.unmount();
        await Promise.resolve();
      });
    }
    root = undefined;
    host.remove();
    vi.unstubAllGlobals();
  });

  it("连续点击合成大文档仍会重新创建预览", async () => {
    await act(async () => {
      root = createRoot(host);
      root.render(createElement(WordPreviewDemo));
      await Promise.resolve();
    });

    const button = Array.from(host.querySelectorAll("button")).find((b) =>
      b.textContent?.includes("合成大文档"),
    );
    expect(button).toBeDefined();

    await act(async () => {
      button!.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await Promise.resolve();
    });
    await waitFor(() => createWordPreviewMock.mock.calls.length === 1);

    await act(async () => {
      button!.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await Promise.resolve();
    });
    await waitFor(() => createWordPreviewMock.mock.calls.length === 2);
    // 第二次必须销毁前一个实例，而不是叠挂
    const first = createWordPreviewMock.mock.results[0]!.value as {
      destroy: ReturnType<typeof vi.fn>;
    };
    expect(first.destroy).toHaveBeenCalled();
  });
});
