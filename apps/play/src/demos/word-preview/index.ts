import { registerDemo } from "../../common/registry.js";

/** 侧边栏父菜单名：word-preview 的 demo 聚合。 */
const GROUP = "word-preview";

registerDemo({
  name: "word-preview",
  category: "preview", // 大类：文档预览
  label: "word-preview — docx 只读预览",
  menuLabel: "文件预览",
  group: GROUP,
  description:
    "上传 / 拖入 .docx 文件进行视觉保真预览（docx-preview 引擎）：格式嗅探、错误归一、超链接 scheme 白名单、多实例样式隔离；附合成大文档性能测量与双实例演示。",
  // 实现按需加载：进入该 demo 才 import 预览器（引擎层较重，首页保持轻量）。
  async load() {
    return import("./word.demo.js");
  },
});
