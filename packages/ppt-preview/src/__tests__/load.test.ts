/**
 * 真实解析矩阵：直接驱动上游官方 Node 入口（@silurus/ooxml/node）验证
 * 本包 fixture 与错误形状假设，并把真实错误接入 normalizeError 做端到端
 * 归一断言（比纯构造的错误对象更强的证据）。
 *
 * Node 入口的能力边界（docs/ppt-preview-spike-record.md S3）：走 Rust/WASM
 * 解析路径，**不含 cfb-guard、不支持 OLE2/加密**——加密样本在 Node 侧的
 * 行为是 not-ooxml，真实解密验收在浏览器（S2 spike 已完成首次，阶段 4
 * 复验）。浏览器入口（PptxPresentation）的加密路径单测在 render/lifecycle
 * 里用 mock 覆盖归一逻辑。
 */
import { describe, expect, it } from "vitest";
import { openPptxPresentation } from "@silurus/ooxml/node";
import { normalizeError } from "../errors";
import { encryptAgile } from "./agile-crypto";
import {
  buildPptx,
  docxBytes,
  fakeOle2Bytes,
  plainZipBytes,
  simplePptx,
} from "./fixtures";

function errOf(
  p: Promise<unknown>,
): Promise<{ code?: string; message: string }> {
  return p.then(
    () => {
      throw new Error("expected rejection");
    },
    (e: unknown) => {
      const err = e as { code?: string; message?: string };
      return { code: err?.code, message: String(err?.message) };
    },
  );
}

describe("官方 Node 入口：真实解析矩阵", () => {
  it("正常 deck：open/slideCount/slideSize/close 全链路", async () => {
    const session = await openPptxPresentation(
      buildPptx({ texts: ["一", "二", "三"] }),
    );
    expect(session.slideCount).toBe(3);
    expect(session.slideWidth).toBe(12192000);
    expect(session.slideHeight).toBe(6858000);
    await session.close();
  });

  it("自定义尺寸（4:3）透传", async () => {
    const session = await openPptxPresentation(
      buildPptx({ texts: ["a"], size: { cx: 9144000, cy: 6858000 } }),
    );
    expect(session.slideWidth).toBe(9144000);
    expect(session.slideHeight).toBe(6858000);
    await session.close();
  });

  it("空 deck 被接受：slideCount=0（本包空 deck 语义的上游依据）", async () => {
    const session = await openPptxPresentation(buildPptx({ texts: [] }));
    expect(session.slideCount).toBe(0);
    await session.close();
  });

  it.each([
    ["plain zip", () => plainZipBytes()],
    ["跨格式 docx", () => docxBytes()],
    ["截断 pptx", () => simplePptx().slice(0, 200)],
    ["随机字节", () => new Uint8Array(512).map((_, i) => (i * 7) % 256)],
    ["空输入", () => new Uint8Array(0)],
    ["假 OLE2（魔数无 CFB 目录）", () => fakeOle2Bytes()],
  ] as const)(
    "%s → not-ooxml 且经 normalizeError 归 UNSUPPORTED",
    async (_label, make) => {
      const err = await errOf(openPptxPresentation(make()));
      expect(err.code).toBe("not-ooxml");
      const normalized = normalizeError(
        Object.assign(new Error(err.message), { code: err.code }),
      );
      expect(normalized.code).toBe("UNSUPPORTED");
    },
  );

  it("docx 的 message 区分跨格式（缺 ppt/presentation.xml）", async () => {
    const err = await errOf(openPptxPresentation(docxBytes()));
    expect(err.message).toContain("ppt/presentation.xml");
  });

  it("加密样本（真实 Agile 容器）：Node 入口不支持 OLE2 → not-ooxml（已核实边界，非 bug）", async () => {
    const enc = encryptAgile(simplePptx(), "pw", { spinCount: 10 });
    // 冒烟加密生成器产出：至少是结构合法的 CFB（魔数 + 足够长度）
    expect(enc[0]).toBe(0xd0);
    expect(enc.length).toBeGreaterThan(4096);
    const err = await errOf(openPptxPresentation(enc));
    expect(err.code).toBe("not-ooxml");
    // 本包 sniff 放行 OLE2（引擎分流），归一后是 UNSUPPORTED
    const normalized = normalizeError(
      Object.assign(new Error(err.message), { code: err.code }),
    );
    expect(normalized.code).toBe("UNSUPPORTED");
  });

  it("缺件 deck（删 presentation.xml）→ not-ooxml", async () => {
    const bytes = buildPptx({ texts: ["a"], omit: ["ppt/presentation.xml"] });
    const err = await errOf(openPptxPresentation(bytes));
    expect(err.code).toBe("not-ooxml");
  });
});
