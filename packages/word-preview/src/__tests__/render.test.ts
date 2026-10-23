// @vitest-environment happy-dom
/**
 * createWordPreview 渲染路径测试（happy-dom）。
 *
 * 能力边界（计划第七轮实证）：happy-dom 的 DOMParser 对带前缀属性保留
 * 完整 localName（r:id / w:val 读不到），只覆盖**文本/结构**路径——段落、
 * span、style 注入、section 数、destroy 竞态；样式 ID、编号、图片、超链接
 * 依赖 r:id 的特性在 happy-dom 下不可靠，安全类断言见 security.test.ts
 * （mock / 非标准 id fixture），真实浏览器复核在 play/docs demo。
 */
import { describe, expect, it, vi } from "vitest";
import { createWordPreview } from "../index";
import type { WordPreviewError } from "../types";
import {
  buildDocx,
  simpleDocx,
  fakeOle2Bytes,
  pageBreakBetween,
  plainZipBytes,
} from "./fixtures";

/**
 * 状态轮询等待（替代固定 sleep 的紧时间窗口——turbo 多任务并行时负载
 * 抖动会让固定 20/50ms 偶发不够）：10ms 间隔轮询条件，默认 2s 超时。
 */
async function waitFor(cond: () => boolean, timeoutMs = 2000): Promise<void> {
  const t0 = Date.now();
  while (!cond()) {
    if (Date.now() - t0 > timeoutMs)
      throw new Error("waitFor: condition not met in time");
    await new Promise((res) => setTimeout(res, 10));
  }
}

/**
 * destroy 静默路径没有正向信号（回调按设计静默），只能给足窗口：真实渲染
 * 实测 5–60ms，150ms 有 2.5–30 倍余量（此前 50ms 在高负载下出过一次抖动）。
 */
async function flushBoot(): Promise<void> {
  await new Promise((res) => setTimeout(res, 150));
}

function makeContainer(): HTMLElement {
  document.body.textContent = "";
  const c = document.createElement("div");
  document.body.appendChild(c);
  return c;
}

describe("createWordPreview 渲染闭环", () => {
  it("最小 docx：段落文本渲染进容器，onRendered 报 pageCount", async () => {
    const container = makeContainer();
    const onRendered = vi.fn();
    createWordPreview(container, {
      source: simpleDocx("Hello 麦克风 world"),
      onRendered,
    });
    await waitFor(() => onRendered.mock.calls.length > 0);

    expect(container.textContent).toContain("Hello 麦克风 world");
    expect(container.querySelector("p")).toBeTruthy();
    expect(container.querySelector("span")).toBeTruthy();
    // 上游注入 <style>（默认样式 + 主题变量）
    expect(container.querySelector("style")).toBeTruthy();
    // 单 section（一页）
    expect(onRendered).toHaveBeenCalledTimes(1);
    expect(onRendered.mock.calls[0][0].pageCount).toBe(1);
    expect(onRendered.mock.calls[0][0].durationMs).toBeGreaterThanOrEqual(0);
  });

  it("显式分页符：pageCount = 2（两个 section 元素）", async () => {
    const container = makeContainer();
    const onRendered = vi.fn();
    createWordPreview(container, {
      source: buildDocx({ body: pageBreakBetween("page one", "page two") }),
      onRendered,
    });
    await waitFor(() => onRendered.mock.calls.length > 0);

    expect(container.querySelectorAll("section").length).toBe(2);
    expect(onRendered.mock.calls[0][0].pageCount).toBe(2);
  });

  it("错误归一：zip 非 docx → UNSUPPORTED（不透上游模糊错）", async () => {
    const container = makeContainer();
    const onError = vi.fn();
    createWordPreview(container, { source: plainZipBytes(), onError });
    await waitFor(() => onError.mock.calls.length > 0);

    expect(onError).toHaveBeenCalledTimes(1);
    const err = onError.mock.calls[0][0] as WordPreviewError;
    expect(err.code).toBe("UNSUPPORTED");
    expect(err.message).toContain("word/document.xml");
    expect(container.textContent).toBe("");
    // 失败即无副作用：私有 wrapper 不得残留在用户容器（空 div 不影响
    // textContent，须以 children 断言——真实浏览器验收发现的原缺陷）
    expect(container.children.length).toBe(0);
  });

  it("错误归一：OLE2 → LEGACY_FORMAT（合并表述）", async () => {
    const container = makeContainer();
    const onError = vi.fn();
    createWordPreview(container, { source: fakeOle2Bytes(), onError });
    await waitFor(() => onError.mock.calls.length > 0);

    expect(onError).toHaveBeenCalledTimes(1);
    expect((onError.mock.calls[0][0] as WordPreviewError).code).toBe(
      "LEGACY_FORMAT",
    );
    expect(container.children.length).toBe(0);
  });

  it("内部唯一 className：实例类名不是上游默认的 docx", async () => {
    const container = makeContainer();
    createWordPreview(container, { source: simpleDocx("cls") });
    await waitFor(() => container.querySelector("style") !== null);

    const wrapper = container.firstElementChild as HTMLElement;
    expect(wrapper).toBeTruthy();
    expect(wrapper.className).toMatch(/^docx-\d+$/);
    expect(wrapper.className).not.toBe("docx");
  });

  it("destroy 竞态：create → 立即 destroy → 容器保持干净、无回调", async () => {
    const container = makeContainer();
    const onRendered = vi.fn();
    const onError = vi.fn();
    const inst = createWordPreview(container, {
      source: simpleDocx("late arrival"),
      onRendered,
      onError,
    });
    inst.destroy();
    // destroy 后等宏任务：renderAsync 迟到完成（上游在 await 之后才会
    // 清空并 append——必须写入已摘除的私有 wrapper，不得回填用户容器）
    await flushBoot();

    expect(container.textContent).toBe("");
    expect(container.children.length).toBe(0);
    expect(onRendered).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
  });

  it("destroy 后同容器复用新实例：新实例内容不被旧 destroy 误删", async () => {
    const container = makeContainer();
    const old = createWordPreview(container, { source: simpleDocx("old") });
    old.destroy();
    const onRendered = vi.fn();
    createWordPreview(container, { source: simpleDocx("new"), onRendered });
    await waitFor(() => onRendered.mock.calls.length > 0);

    expect(container.textContent).toContain("new");
    expect(onRendered).toHaveBeenCalledTimes(1);
  });

  it("destroy 幂等：二次 destroy 不抛错", () => {
    const container = makeContainer();
    const inst = createWordPreview(container, { source: simpleDocx("x") });
    inst.destroy();
    inst.destroy();
  });
});
