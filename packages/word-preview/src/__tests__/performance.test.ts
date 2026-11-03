// @vitest-environment happy-dom
import { describe, it, expect, beforeAll } from "vitest";
import { renderAsync } from "docx-preview";
import { createWordPreview } from "../index";
import { bigDocxBytes } from "./fixtures";

// 与 excel-exporter / excel-preview 同一模式：本地阈值即产品 SLA
// （PERF_SLACK 缺省为 1）；PR CI 设 RUN_PERF=0 跳过，release 链路
// （release.yml）以显式宽限因子（PERF_SLACK）跑性能门——CI 共享 runner
// 上 DOM 构建实测慢 ~3x（2026-10-24：446ms vs 本地中位 ~96ms），无宽限
// 必挂（此前遗漏接入导致发版被卡）。机器高负载偶发失败先复跑排除抖动。
const SLACK = Number(process.env.PERF_SLACK ?? "1") || 1;
const RUN_PERF = process.env.RUN_PERF !== "0";

// 口径：happy-dom 度量的是**DOM 构建**耗时（1000 段落 + 100×5 表格的纯
// 文本文档、无图片/内嵌字体——resolve 前 await 全部资源任务，fixture 放
// 图会把 base64 编码耗时计入）。happy-dom 无真实布局/绘制，不能代表真实
// 渲染性能；真实 Chromium 的完整渲染基线（含资源加载）在 README/limits
// 以实测数据另行标注。
// 阈值定档依据（2026-10-23 本机复测 6 轮 + turbo 并行 1 轮）：本测试是
// 单次冷调用形态（含 JIT/首次解析开销），实测中位 ~96ms、最大 99ms、
// 并行 99ms；首轮定档的"中位 63ms"系预热态测量、与测试形态不符（已订正）。
// 150ms 留 ~55% 负载余量，与 excel-exporter 两基线 27-57% 的余量同档。
describe.runIf(RUN_PERF)(
  "performance (DOM build, 1000 paras + 100x5 table)",
  () => {
    let container: HTMLElement;

    beforeAll(() => {
      document.body.textContent = "";
      container = document.createElement("div");
      document.body.appendChild(container);
    });

    it("renderAsync < 150ms", async () => {
      const bytes = bigDocxBytes(1000, 100, 5);
      const t0 = performance.now();
      await renderAsync(bytes, container, undefined, {
        // 与 createWordPreview 交给上游引擎的默认渲染参数一致；本基线直接
        // 调用引擎，不包含本包的格式嗅探与第二次 JSZip 目录解析开销
        className: "docx-perf",
        useBase64URL: true,
        ignoreLastRenderedPageBreak: false,
        renderAltChunks: false,
      });
      const dt = performance.now() - t0;
      // 产物规模断言（防止 fixture 退化后空跑出假绿）
      expect(container.querySelectorAll("p").length).toBe(1500);
      expect(container.querySelectorAll("tr").length).toBe(100);
      expect(dt).toBeLessThan(150 * SLACK);
    });

    // 全链路门：createWordPreview → onRendered（源读取 + 嗅探的 JSZip
    // 中央目录解析 + [Content_Types].xml/_rels 解压 + 引擎渲染 + 安全
    // 后处理）。引擎直调基线不覆盖嗅探链路——嗅探若引入回归（如部件
    // 查找复杂度恶化）此前无任何门可拦。阈值拦数量级回归，不追精确：
    // 引擎基线 150ms + 嗅探/二次解析/后处理实测 <100ms + 负载余量。
    it("createWordPreview full pipeline < 400ms", async () => {
      container.textContent = ""; // 隔离上一用例的 DOM 产物
      const bytes = bigDocxBytes(1000, 100, 5);
      let rendered: (() => void) | null = null;
      const renderedPromise = new Promise<void>((res) => {
        rendered = res;
      });
      const t0 = performance.now();
      createWordPreview(container, {
        source: bytes,
        onRendered: () => rendered?.(),
      });
      await renderedPromise;
      const dt = performance.now() - t0;
      expect(container.querySelectorAll("p").length).toBeGreaterThanOrEqual(
        1500,
      );
      expect(container.querySelectorAll("tr").length).toBeGreaterThanOrEqual(
        100,
      );
      expect(dt).toBeLessThan(400 * SLACK);
    });
  },
);
