import { registerDemo } from "../../common/registry.js";

/** 侧边栏父菜单名：excel-preview 的 demo 聚合。 */
const GROUP = "excel-preview";

registerDemo({
  name: "excel-preview",
  category: "preview", // 大类：文档预览
  label: "excel-preview — xlsx 只读预览",
  menuLabel: "文件预览",
  group: GROUP,
  description:
    "上传 / 拖入 xlsx / xlsm / 加密 / CSV 文件，Worker 解析 + DOM 虚拟滚动渲染：样式（含主题色）、合并、冻结、数字格式（Excel 实际行为）。",
  // 实现按需加载：进入该 demo 才 import 预览器与 WASM/worker 资产。
  async load() {
    return import("./preview.demo.js");
  },
});
