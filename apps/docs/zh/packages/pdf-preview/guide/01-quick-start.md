# 快速上手

前提：浏览器 DOM。import 本包是 SSR 安全的（模块顶层不求值任何 DOM 全局——已对 pdfjs-dist 6.4.29 实测验证），但 `createPdfPreview` 本身必须在浏览器中调用。

## 1. 安装

```bash
pnpm add @marcusok/pdf-preview
```

全部随包分发：pdf.js 引擎打进 `dist/index.js`，worker 打进 `dist/pdf.worker.min.mjs`，CJK/标准字体资产打进 `dist/cmaps/` + `dist/standard_fonts/`（载荷合计约 4.5 MB）。零传递依赖。

## 2. 挂载预览

```ts
import { createPdfPreview } from "@marcusok/pdf-preview";

const preview = createPdfPreview(container, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onRendered: ({ pageCount, firstPageMs, durationMs }) => {
    console.log(`共 ${pageCount} 页；首页出图 ${firstPageMs}ms`);
  },
  onError: (e) => {
    // e.code: UNSUPPORTED | PASSWORD_PROTECTED | CORRUPT | UNKNOWN
    console.error(e.code, e.message);
  },
});
```

页面进入视口才渲染——`onRendered` 在**首页**绘制完成时触发（`pageCount` 是文档总页数）。懒渲染管线不存在"全部页渲染完成"事件，这是有意设计。

## 3. 卸载

```ts
preview.destroy();
```

`destroy()` 断开渲染观察器、取消在途渲染任务、终止 pdf.js worker 并同步摘除 DOM。迟到的回调静默（取消不是错误）。可重复调用；容器可立即复用。

## 资产接入——唯一必须做对的事

worker 通过 `new URL("./pdf.worker.min.mjs", import.meta.url)` 字面量定位。Vite（8+）、webpack 5 等管道的**生产构建**会改写该表达式并把 worker emit 为带 base 的哈希资产——build 模式零配置；直接以 ESM/CDN 引用包 `dist` 时也天然正确。**Vite 5–7 的 dev server 不同**——见下方注意事项。

`cmaps/` 与 `standard_fonts/` 是**目录**——bundler 不会复制目录。分场景：

| 场景                                            | Worker                                                                                         | cmaps / standard_fonts                                                                              |
| ----------------------------------------------- | ---------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| 直接 ESM / CDN 引用包                           | 从 `dist` 解析                                                                                 | 从 worker URL 推导同层目录——开箱即用                                                                |
| Vite 8+ / webpack 5 默认                        | 自动 emit ✅                                                                                   | 推导 URL 落在哈希资产目录内 → **404、中文静默劣化** ❌                                              |
| Vite 5–7 dev server（npm 安装）                 | 预打包副本指向 `.vite/deps` → **404** ❌                                                       | 推导 URL 因同一原因失效 ❌                                                                          |
| Vite + `vite-static-copy` / webpack + copy 插件 | 自动 emit ✅                                                                                   | 把 `dist/cmaps` + `dist/standard_fonts` 拷到 public 目录，显式传 `cMapUrl`/`standardFontDataUrl` ✅ |
| 自托管（自己提供文件）                          | 真实可访问 URL——如 `import workerSrc from "@marcusok/pdf-preview/dist/pdf.worker.min.mjs?url"` | 提供 `dist/cmaps/` + `dist/standard_fonts/` 的 URL 并显式传入 ✅                                    |

> **Vite 5–7 dev server（依赖预打包）。** Vite 5–7 用 esbuild 把 npm 安装的依赖预打包进 `node_modules/.vite/deps/`。预打包后的副本会原样保留 `new URL("./pdf.worker.min.mjs", import.meta.url)`，于是解析到 `.vite/deps/pdf.worker.min.mjs`——该文件并不存在（Vite 用 HTML fallback 或 404 响应）。pdf.js 无法构造/导入 worker，预览报 `UNKNOWN`（"Failed to load the PDF worker"）。把包加入 `optimizeDeps.exclude` 并重启 dev server：
>
> ```ts
> // vite.config.ts
> export default { optimizeDeps: { exclude: ["@marcusok/pdf-preview"] } };
> ```
>
> 排除后 Vite 直接从 `node_modules` 提供包的 ESM 产物与资产，`import.meta.url` 重新指向真实 `dist/`。生产构建不受影响。（Vite 8 的 rolldown 预打包会正确改写该字面量；文档站使用 workspace 链接包、不参与预打包，因此不会走到 npm 安装形态。）

推荐的 bundler 配置（Vite）：

```ts
// vite.config.ts
import { viteStaticCopy } from "vite-plugin-static-copy";

export default {
  plugins: [
    viteStaticCopy({
      targets: [
        {
          src: "node_modules/@marcusok/pdf-preview/dist/cmaps/*",
          dest: "pdf/cmaps",
        },
        {
          src: "node_modules/@marcusok/pdf-preview/dist/standard_fonts/*",
          dest: "pdf/standard_fonts",
        },
      ],
    }),
  ],
};
```

```ts
createPdfPreview(el, {
  source,
  cMapUrl: "/pdf/cmaps/",
  standardFontDataUrl: "/pdf/standard_fonts/",
});
```

如果应用部署在子路径下，请给这些 URL 加上 base（`import.meta.env.BASE_URL` 或 VitePress 的 `withBase`）：根绝对路径 `/pdf/cmaps/` 只在域名根目录下正确。

cmaps 缺失**不报错**——CJK 文档直接渲染出空白/乱码文本，这正是这步配置如此关键的原因。用真实中文 PDF 做一次冒烟验证是唯一可靠的检查方式。

## 缩放

```ts
createPdfPreview(el, { source, scale: "fit-width" }); // 默认
createPdfPreview(el, { source, scale: "fit-page" });
createPdfPreview(el, { source, scale: 1.5 }); // 原始 viewport 缩放
```

`fit-*` 在创建时对容器量测一次。容器隐藏（0 宽）退回 `scale = 1` 并 `console.warn`。容器尺寸变化不自动重拟合——destroy 后重建实例。

## 下一步

- [限制](./02-limits)——内存行为、wasm 门控特性、CSP 要求、SSR 边界。
- 在 [play](/zh/play) 用你自己的文件试用。
