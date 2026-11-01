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

/** 以给定 File 触发 change（连续选择场景需要不同的文件对象）。 */
function selectFile(host: HTMLElement, file: File): void {
  const input = host.querySelector<HTMLInputElement>('input[type="file"]')!;
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

/** 可控读取时长的文件：arrayBuffer 由测试显式 resolve。 */
function deferredFile(
  name: string,
  bytes: Uint8Array,
): {
  file: File;
  resolve: () => void;
} {
  const file = new File([new Uint8Array(bytes).buffer], name);
  let release!: (buf: ArrayBuffer) => void;
  file.arrayBuffer = () =>
    new Promise<ArrayBuffer>((resolve) => {
      release = resolve;
    });
  return {
    file,
    resolve: () => release(new Uint8Array(bytes).buffer),
  };
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

  it("快速连续选择不同文件时只采用最后一次选择", async () => {
    await act(async () => {
      root = createRoot(host);
      root.render(createElement(ExcelPreviewDemo));
      await Promise.resolve();
    });

    // 先选的大文件读取故意挂起，后选的小文件先读完——旧实现会让大文件
    // 在挂载小文件之后才写入 bytesRef/renderSeq，旧数据覆盖新选择。
    const slow = deferredFile("large.xlsx", new Uint8Array([1]));
    const small = new File([new Uint8Array([2]).buffer], "small.xlsx");

    await act(async () => {
      selectFile(host, slow.file);
      selectFile(host, small);
      await Promise.resolve();
      await Promise.resolve();
    });
    await waitFor(() => createPreviewMock.mock.calls.length >= 1);

    await act(async () => {
      slow.resolve();
      await Promise.resolve();
      await Promise.resolve();
    });
    // 给迟到的读取结果一个状态更新窗口（旧实现会在这里创建第二个实例）
    await new Promise((resolve) => setTimeout(resolve, 50));

    expect(createPreviewMock.mock.calls).toHaveLength(1);
    const source = createPreviewMock.mock.calls[0]![1].source as Uint8Array;
    expect(Array.from(source)).toEqual([2]);
  });
});
