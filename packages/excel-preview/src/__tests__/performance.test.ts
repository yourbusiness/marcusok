// @vitest-environment happy-dom
import { describe, it, expect, beforeAll } from "vitest";
import { readBuffer } from "@marcusok/xlsx-core";
import { buildPreviewWorkbook } from "../parse/model";
import { SheetRenderer } from "../render/renderer";
import { buildRowIndex, forEachCellInRange } from "../render/cell-index";
import type { PreviewCell } from "../types";
import { bigWorkbookBytes } from "./fixtures";

// 与 excel-exporter 同一模式：本地阈值即产品 SLA（PERF_SLACK 缺省为 1）；
// PR CI 设 RUN_PERF=0 跳过，release 链路（release.yml）以显式宽限因子
// （PERF_SLACK）跑性能门。机器高负载偶发失败先复跑排除抖动。
const SLACK = Number(process.env.PERF_SLACK ?? "1") || 1;
const RUN_PERF = process.env.RUN_PERF !== "0";

// 口径（规划第七节第 6 条）：vitest 只测"解析 + 视图模型构建 + 首屏 DOM
// 节点创建"。happy-dom 无真实布局，不能度量渲染耗时；真实渲染性能在
// play/浏览器实测并单独标注口径。
describe.runIf(RUN_PERF)(
  "performance (parse + model + first-screen DOM)",
  () => {
    let container: HTMLElement;

    beforeAll(() => {
      document.body.textContent = "";
      container = document.createElement("div");
      document.body.appendChild(container);
    });

    it("parse 10k rows x 8 cols (model build) < 400ms", async () => {
      const bytes = bigWorkbookBytes(10_000, 8);
      const t0 = performance.now();
      const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
      const dt = performance.now() - t0;
      expect(model.sheets[0].rowCount).toBe(10_000);
      expect(dt).toBeLessThan(400 * SLACK);
    });

    it("parse 100k rows x 8 cols (model build) < 4s", async () => {
      const bytes = bigWorkbookBytes(100_000, 8);
      const t0 = performance.now();
      const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
      const dt = performance.now() - t0;
      expect(model.sheets[0].rowCount).toBe(100_000);
      expect(dt).toBeLessThan(4000 * SLACK);
    });

    it("first-screen DOM build (10k sheet) < 150ms", async () => {
      const bytes = bigWorkbookBytes(10_000, 8);
      const model = buildPreviewWorkbook(await readBuffer(bytes), bytes);
      const t0 = performance.now();
      const r = new SheetRenderer(container);
      r.render(model, {});
      const dt = performance.now() - t0;
      expect(container.querySelectorAll(".xpv-cell").length).toBeGreaterThan(
        50,
      );
      expect(dt).toBeLessThan(150 * SLACK);
      r.destroy();
    });

    it("wide-row window scan (60 rows x 16384 cols, 200 probes) < 30ms", () => {
      // 宽表横向滚动的每帧成本守卫：窗口差量循环经 cell-index 二分后，
      // 每次探测只触碰窗口内的格子。旧形态（per-row Map 全量遍历再按列
      // 过滤）在同一数据上本机实测 ~430ms（67x），回归会在本阈值爆红。
      const COLS = 16_384;
      const ROWS = 60;
      const WIN = 40;
      const rows = Array.from({ length: ROWS }, (_, r) => ({
        index: r + 1,
        height: null,
        hidden: false,
        cells: Array.from({ length: COLS }, (_, c): PreviewCell => ({
          col: c,
          type: "number" as const,
          value: String(r * COLS + c),
          styleIndex: null,
        })),
      }));
      const index = buildRowIndex(rows);
      const t0 = performance.now();
      let hits = 0;
      for (let k = 0; k < 200; k++) {
        // 7919 与 16000 互素，起点在列轴上充分散布
        const colStart = (k * 7919) % (COLS - WIN);
        for (let r = 0; r < ROWS; r++) {
          forEachCellInRange(index.get(r), colStart, colStart + WIN, () => {
            hits++;
          });
        }
      }
      const dt = performance.now() - t0;
      expect(hits).toBe(200 * ROWS * WIN); // 规模断言防空跑
      expect(dt).toBeLessThan(30 * SLACK);
    });
  },
);
