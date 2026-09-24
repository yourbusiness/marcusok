import { registerDemo } from "../../common/registry.js";

// 文档预览类的预留占位：该分类暂无交付包，先注册一个空白页面把
// 侧边栏与首页的"文档预览"分区建出来；包落地后整个目录由真实
// demo 替换（不设 group，一级菜单直达占位页）。
registerDemo({
  name: "doc-preview",
  category: "preview",
  label: "doc-preview — 文档预览（预留）",
  // 分类标题已标明"文档预览"，菜单项用短名即可
  menuLabel: "doc-preview",
  description:
    "文档预览类的预留占位：分类已建、包未落地，页面暂为空白，交付后由真实 demo 替换。",
  async load() {
    return import("./placeholder.demo.js");
  },
});
