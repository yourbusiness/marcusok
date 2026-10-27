import { registerDemo } from "../../common/registry.js";

/** 侧边栏父菜单名：pdf-preview 的 demo 聚合。 */
const GROUP = "pdf-preview";

registerDemo({
  name: "pdf-preview",
  category: "preview", // 大类：文档预览
  label: "pdf-preview — PDF 只读预览",
  menuLabel: "文件预览",
  group: GROUP,
  description:
    "上传 / 拖入 .pdf 文件进行连续滚动懒渲染预览（pdf.js 引擎）：视口外页自动释放、文本层可选中复制、fit-width/fit-page/数值缩放；附 200 页合成大文档与双实例演示。中文 PDF 走 cmaps 链路，可上传真实中文文档验证。",
  // 实现按需加载：进入该 demo 才 import 预览器（pdf.js 主 bundle 约 850KB）。
  async load() {
    return import("./pdf.demo.js");
  },
});
