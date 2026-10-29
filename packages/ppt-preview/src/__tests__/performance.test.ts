/**
 * 性能基线（node + 官方 Node 入口直测，真实解析、禁 mock）。
 *
 * 口径：200 页合成 deck 的**稳态**全量解析（openPptxPresentation →
 * slideCount → close）。Rust/WASM 解析极快（2026-10-29 本机实测 5 轮
 * 中位 3ms），50ms 阈值留 ~15 倍余量覆盖 CI 共享 runner 波动。
 *
 * **这个数字不代表浏览器首屏**：浏览器路径另有 module Worker 启动、
 * wasm 加载、全 deck preflight 与主线程 Canvas 渲染/文本层（加密文件
 * 还叠加主线程 SHA-512 密钥派生，约 +1s）——首屏指标在真实 Chromium
 * 验收并记入 README/limits（pdf-preview 的双基线同思路：Node 管解析
 * 回归，浏览器管首屏 SLA）。
 *
 * 样本规模依据：非渐进 load 在 resolve 前会按页序 preflight 全部页
 * （docs/ spike 记录事实 15），200 页足以暴露随页数线性增长的成本；
 * 产物规模断言防 fixture 空跑。
 */
import { beforeAll, describe, expect, it } from "vitest";
import { openPptxPresentation } from "@silurus/ooxml/node";
import { buildPptx } from "./fixtures";

// 仓库各发布包统一口径：本地阈值即产品 SLA（PERF_SLACK 缺省 1）；CI /
// release / pre-push 设 RUN_PERF=0 跳过；release.yml 以 PERF_SLACK=4 给
// 共享 runner 宽限。机器高负载偶发失败先复跑排除抖动。
const SLACK = Number(process.env.PERF_SLACK ?? "1") || 1;
const RUN_PERF = process.env.RUN_PERF !== "0";

/** 200 页稳态全量解析阈值（ms）。 */
const BASELINE_MS = 50;
const SLIDES = 200;
/** 预热 1 次 + 采样 3 次，取中位数。 */
const SAMPLES = 3;

describe.runIf(RUN_PERF)(
  "performance (full-deck parse, 200 slides, node)",
  () => {
    let bytes: Uint8Array;

    /**
     * 跑一次完整解析并返回耗时。每次用独立字节副本并 close session
     * （WASM 侧内存随 session 生命周期回收，复用同一份 source 也会让
     * 跨实现的行为不可控——与包内 toBytes 的复制理由相同）。
     */
    async function measureOnce(): Promise<number> {
      const t0 = performance.now();
      const session = await openPptxPresentation(bytes.slice());
      try {
        expect(session.slideCount).toBe(SLIDES); // 产物规模断言防 fixture 空跑
        expect(session.slideWidth).toBe(12192000);
        return performance.now() - t0;
      } finally {
        await session.close();
      }
    }

    beforeAll(() => {
      bytes = buildPptx({
        texts: Array.from(
          { length: SLIDES },
          (_, i) => `第${i + 1}页 性能基线样本`,
        ),
      });
      // 预热：只付一次 WASM 模块实例化成本，不计入稳态 SLA
      return measureOnce();
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
