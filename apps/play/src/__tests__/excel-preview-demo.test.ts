// @vitest-environment happy-dom
import { act } from "react";
import { createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const { createPreviewMock } = vi.hoisted(() => ({
  createPreviewMock: vi.fn(),
}));

vi.mock("@marcusok/excel-preview", () => ({
  createPreview: createPreviewMock,
}));

import ExcelPreviewDemo from "../demos/excel-preview/preview.demo";

async function waitFor(cond: () => boolean, timeoutMs = 2000): Promise<void> {
  const t0 = Date.now();
  while (!cond()) {
    if (Date.now() - t0 > timeoutMs) {
      throw new Error("waitFor: condition not met in time");
    }
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

async function upload(host: HTMLElement, bytes: Uint8Array): Promise<void> {
  const input = host.querySelector<HTMLInputElement>('input[type="file"]')!;
  const copy = new Uint8Array(bytes);
  const file = new File([copy.buffer], "same-name.xlsx");
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  await act(async () => {
    input.dispatchEvent(new Event("change", { bubbles: true }));
    await Promise.resolve();
  });
}

describe("excel-preview demo", () => {
  let root: Root | undefined;
  let host: HTMLElement;

  beforeEach(() => {
    createPreviewMock.mockReset();
    createPreviewMock.mockImplementation(() => ({
      destroy: vi.fn(),
      setSheet: vi.fn(),
      getSheetNames: vi.fn(() => []),
    }));
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

  it("同名文件再次上传且内容变化时重新创建预览", async () => {
    await act(async () => {
      root = createRoot(host);
      root.render(createElement(ExcelPreviewDemo));
      await Promise.resolve();
    });

    await upload(host, new Uint8Array([1]));
    await waitFor(() => createPreviewMock.mock.calls.length === 1);
    await upload(host, new Uint8Array([2]));
    await waitFor(() => createPreviewMock.mock.calls.length === 2);

    const secondSource = createPreviewMock.mock.calls[1]![1]
      .source as Uint8Array;
    expect(Array.from(secondSource)).toEqual([2]);
  });
});
