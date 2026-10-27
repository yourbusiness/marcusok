/**
 * 性能基线（node + legacy 引擎直测）。
 *
 * 口径：200 页纯文本 PDF 的**稳态**解析 + 首页文本提取全程
 * （getDocument → getPage(1) → streamTextContent 读尽）——这是懒渲染
 * 形态下首屏出图前的主线程外准备工作量。canvas 渲染与文本层排版在
 * 真实浏览器（play/docs demo）验收，happy-dom 无真实绘制不代表渲染
 * 性能（word-preview 双基线同思路）。
 *
 * 测量形态：先预热一次，再取 3 次运行的中位数。pdf.js 首次 getDocument
 * 会初始化 fake worker / 加载 worker 模块（本机 ~57ms、CI 首次 ~540ms），
 * 这是进程级的一次性成本，不属于每次解析的 SLA；README 的 "~66ms 中位"
 * 也是热态口径。2026-10-27 release 失败复盘：旧测试只测冷启动，CI 实测
 * 639ms > 150×4=600ms；修正的是测量口径，阈值与 PERF_SLACK 均未放宽。
 */
import { describe, it, expect, beforeAll } from "vitest";
import { getDocument } from "pdfjs-dist";
import type { PDFDocumentLoadingTask } from "pdfjs-dist";
import { makePdfBytes, STANDARD_FONT_DATA_URL } from "./fixtures";

// 仓库各发布包统一口径：本地阈值即产品 SLA（PERF_SLACK 缺省 1）；CI /
// release / pre-push 设 RUN_PERF=0 跳过；release.yml 以 PERF_SLACK=4 给
// 共享 runner 宽限。机器高负载偶发失败先复跑排除抖动。
const SLACK = Number(process.env.PERF_SLACK ?? "1") || 1;
const RUN_PERF = process.env.RUN_PERF !== "0";

/** 200 页纯文本文档的稳态解析+首页文本提取阈值（ms）。定档依据
 * （2026-10-25 本机 3 轮实测）：中位 ~66ms、最大 70ms——150ms 留 ~55%
 * 负载余量，与 word-preview DOM 基线（96ms 中位 / 150ms 阈值）同档。 */
const BASELINE_MS = 150;

/** 预热 1 次 + 采样 3 次，取中位数。 */
const SAMPLES = 3;

describe.runIf(RUN_PERF)(
  "performance (parse + first-page text extraction, 200 pages)",
  () => {
    let bytes: Uint8Array;

    /**
     * 跑一次完整管线并返回耗时。每次必须用字节副本：getDocument 会
     * transfer `data.buffer`，复用同一份 source 会让下一次读到已 detach
     * 的空字节（与包内 toBytes 的复制原因相同）。
     */
    async function measureOnce(): Promise<number> {
      const t0 = performance.now();
      const task: PDFDocumentLoadingTask = getDocument({
        data: bytes.slice(),
        standardFontDataUrl: STANDARD_FONT_DATA_URL,
      });
      try {
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
        expect(chars).toBeGreaterThan(0);
        return performance.now() - t0;
      } finally {
        await task.destroy();
      }
    }

    beforeAll(async () => {
      bytes = await makePdfBytes({ pages: 200 });
      // 预热：只付一次 pdf.js worker/模块初始化成本，不计入稳态 SLA
      await measureOnce();
    });

    it("median of 3 warmed runs < threshold", async () => {
      const samples: number[] = [];
      for (let i = 0; i < SAMPLES; i++) samples.push(await measureOnce());
      samples.sort((a, b) => a - b);
      const median = samples[Math.floor(SAMPLES / 2)];
      expect(median).toBeLessThan(BASELINE_MS * SLACK);
    });
  },
);
