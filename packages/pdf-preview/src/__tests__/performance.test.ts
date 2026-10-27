/**
 * 性能基线（node + legacy 引擎直测）。
 *
 * 口径：200 页纯文本 PDF 的 **解析 + 首页文本提取** 全程
 * （getDocument → getPage(1) → streamTextContent 读尽）——这是懒渲染
 * 形态下首屏出图前的主线程外准备工作量。canvas 渲染与文本层排版在
 * 真实浏览器（play/docs demo）验收，happy-dom 无真实绘制不代表渲染
 * 性能（word-preview 双基线同思路）。
 */
import { describe, it, expect, beforeAll } from "vitest";
import { getDocument } from "pdfjs-dist";
import type { PDFDocumentLoadingTask } from "pdfjs-dist";
import { makePdfBytes } from "./fixtures";

// 仓库三包统一口径：本地阈值即产品 SLA（PERF_SLACK 缺省 1）；CI / release
// / pre-push 设 RUN_PERF=0 跳过；release.yml 以 PERF_SLACK=4 给共享 runner
// 宽限。机器高负载偶发失败先复跑排除抖动。
const SLACK = Number(process.env.PERF_SLACK ?? "1") || 1;
const RUN_PERF = process.env.RUN_PERF !== "0";

/** 200 页纯文本文档的解析+首页文本提取阈值（ms）。定档依据（2026-10-25
 * 本机 3 轮实测）：中位 ~66ms、最大 70ms——150ms 留 ~55% 负载余量，与
 * word-preview DOM 基线（96ms 中位 / 150ms 阈值）同档。 */
const BASELINE_MS = 150;

describe.runIf(RUN_PERF)(
  "performance (parse + first-page text extraction, 200 pages)",
  () => {
    let bytes: Uint8Array;

    beforeAll(async () => {
      bytes = await makePdfBytes({ pages: 200 });
    });

    it("getDocument + getPage(1) + streamTextContent < threshold", async () => {
      const t0 = performance.now();
      const task: PDFDocumentLoadingTask = getDocument({ data: bytes });
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(200); // 产物规模断言防 fixture 空跑
      const page = await pdf.getPage(1);
      const reader = page.streamTextContent().getReader();
      let chars = 0;
      for (;;) {
        const r = await reader.read();
        if (r.done) break;
        for (const item of r.value.items) {
          if ("str" in item) chars += item.str.length;
        }
      }
      const dt = performance.now() - t0;
      expect(chars).toBeGreaterThan(0);
      expect(dt).toBeLessThan(BASELINE_MS * SLACK);
      await task.destroy();
    });
  },
);
