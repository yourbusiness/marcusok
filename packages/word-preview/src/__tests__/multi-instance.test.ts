// @vitest-environment happy-dom
/**
 * 同页多实例隔离测试（真实上游渲染路径）。
 *
 * 上游默认固定 className "docx"，注入的默认样式/文档样式/编号/主题变量
 * 全部是文档级全局 CSS——双实例会互相覆盖（真实 Chromium 已复现：后渲染
 * 者的编号 ::before 覆盖前者的）。本包为每实例生成唯一 className，
 * 上游生成选择器/变量时整体以该类名为前缀（rootSelector、processStyleName、
 * --docx-N-* 编号变量，源码实证），换名即隔离。
 *
 * happy-dom 边界：不计算 CSS 级联，不做 computed style 断言（真实级联
 * 冲突已在浏览器复现、demo 侧复核）；这里断言结构性事实——类名唯一、
 * 各自 style 只含本实例前缀。
 */
import { describe, expect, it } from "vitest";
import { createWordPreview } from "../index";
import { buildDocx, simpleDocx, parasListXml } from "./fixtures";

/** 状态轮询等待（替代固定 sleep；与 render.test.ts 同款，见其注释）。 */
async function waitFor(cond: () => boolean, timeoutMs = 2000): Promise<void> {
  const t0 = Date.now();
  while (!cond()) {
    if (Date.now() - t0 > timeoutMs)
      throw new Error("waitFor: condition not met in time");
    await new Promise((res) => setTimeout(res, 10));
  }
}

describe("同页多实例样式隔离", () => {
  it("两实例：wrapper 类名互异，各自 style 文本只含本实例类名前缀", async () => {
    document.body.textContent = "";
    const c1 = document.createElement("div");
    const c2 = document.createElement("div");
    document.body.append(c1, c2);

    createWordPreview(c1, {
      source: buildDocx({ body: parasListXml(["AAA"]) }),
    });
    createWordPreview(c2, {
      source: buildDocx({ body: parasListXml(["BBB"]) }),
    });
    await waitFor(
      () =>
        (c1.textContent ?? "").includes("AAA") &&
        (c2.textContent ?? "").includes("BBB"),
    );

    const w1 = c1.firstElementChild as HTMLElement;
    const w2 = c2.firstElementChild as HTMLElement;
    expect(w1.className).toMatch(/^docx-\d+$/);
    expect(w2.className).toMatch(/^docx-\d+$/);
    expect(w1.className).not.toBe(w2.className);

    // 各自渲染容器内的 <style> 文本只含本实例前缀（默认样式的
    // rootSelector .docx-N-wrapper / 主题变量 .docx-N 都会带上）
    const styles1 = [...c1.querySelectorAll("style")]
      .map((s) => s.textContent ?? "")
      .join("\n");
    const styles2 = [...c2.querySelectorAll("style")]
      .map((s) => s.textContent ?? "")
      .join("\n");
    expect(styles1).toContain(`.${w1.className}`);
    expect(styles1).not.toContain(`.${w2.className}`);
    expect(styles2).toContain(`.${w2.className}`);
    expect(styles2).not.toContain(`.${w1.className}`);

    // 两个实例的文本各自落位
    expect(c1.textContent).toContain("AAA");
    expect(c2.textContent).toContain("BBB");
  });

  it("单实例销毁不影响另一实例的 DOM", async () => {
    document.body.textContent = "";
    const c1 = document.createElement("div");
    const c2 = document.createElement("div");
    document.body.append(c1, c2);

    const p1 = createWordPreview(c1, { source: simpleDocx("keep me") });
    createWordPreview(c2, { source: simpleDocx("also keep") });
    await waitFor(() => (c2.textContent ?? "").includes("also keep"));
    p1.destroy();
    // destroy 后给在途链走完的余量（静默路径无正向信号）
    await new Promise((res) => setTimeout(res, 150));

    expect(c1.children.length).toBe(0);
    expect(c2.textContent).toContain("also keep");
  });
});
