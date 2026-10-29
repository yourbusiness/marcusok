import { registerDemo } from "../../common/registry.js";

/** 侧边栏父菜单名：ppt-preview 的 demo 聚合。 */
const GROUP = "ppt-preview";

registerDemo({
  name: "ppt-preview",
  category: "preview", // 大类：文档预览
  label: "ppt-preview — PPTX 只读预览",
  menuLabel: "幻灯片预览",
  group: GROUP,
  description:
    "上传 / 拖入 .pptx 文件进行连续滚动懒渲染预览（Rust/WASM 解析引擎）：视口外页自动释放、文本层可选中复制、0-based 页导航（goToSlide/next/prev）、Agile 加密文档密码解密；附多页合成样本与双实例演示。",
  // 实现按需加载：进入该 demo 才 import 预览器（引擎 chunk 约 2.5MB + wasm 2MB）。
  async load() {
    return import("./ppt.demo.js");
  },
});
