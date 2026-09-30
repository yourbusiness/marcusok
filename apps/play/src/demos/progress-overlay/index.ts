import { registerDemo } from "../../common/registry.js";

/** 侧边栏父菜单名：progress-overlay 的 demo 聚合到同一子菜单下。 */
const GROUP = "progress-overlay";

registerDemo({
  name: "progress-overlay",
  category: "shared", // 大类：公共组件（业务包共用的 UI 层）
  label: "progress-overlay — 通用进度遮罩",
  menuLabel: "遮罩演示",
  group: GROUP,
  description:
    "独立演示 @marcusok/progress-overlay：不确定态旋转圆环 → 流式进度百分比条的切换、主题/文案定制、delayMs 门控。excel-exporter 的 overlay 选项即由它驱动。",
  async load() {
    return import("./basic.demo.js");
  },
});
