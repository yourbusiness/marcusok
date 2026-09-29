// @vitest-environment happy-dom
import { describe, it, expect, beforeAll } from "vitest";
import { readBuffer } from "@marcusok/xlsx-core";
import { buildPreviewWorkbook } from "../parse/model";
import { SheetRenderer } from "../render/renderer";
import { bigWorkbookBytes } from "./fixtures";

// 与 excel-exporter 同一模式：本地阈值即产品 SLA（无环境宽限）；CI 与
// release 链路设 RUN_PERF=0 跳过。机器高负载偶发失败先复跑排除抖动。
const SLACK = 1.0;
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
  },
);
