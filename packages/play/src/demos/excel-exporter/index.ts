import { registerDemo } from "../../common/registry.js";

/** 侧边栏父菜单名：excel-exporter 的全部 demo 聚合到同一子菜单下。 */
const GROUP = "excel-exporter";

registerDemo({
  name: "excel-exporter",
  label: "excel-exporter — Excel 导出引擎",
  menuLabel: "性能对比",
  group: GROUP,
  description:
    "用 mock 数据对比 auto / main / worker / stream 四种导出模式的耗时、吞吐、文件体积与阶段拆解。",
  // 实现按需加载：进入该 demo 时才 import 重型代码（excel-exporter 及其
  // WASM/worker 资产），首页只加载这份轻量元信息。
  async load() {
    return import("./basic-export.demo.js");
  },
});

// 样式演示与性能演示并列：同一个 index.ts 可注册多个 demo（registry 以 name
// 为键），样式预览与导出共用同一份 SheetConfig。
registerDemo({
  name: "excel-exporter-styles",
  label: "excel-exporter — 样式系统演示",
  menuLabel: "样式系统",
  group: GROUP,
  description:
    "StylePresets 内置预设速查 + 三个场景配方（组合使用 / spread 派生 / 合并语义），浏览器内实时预览模拟 Excel 观感，一键导出真实文件验证。",
  async load() {
    return import("./styles.demo.js");
  },
});
