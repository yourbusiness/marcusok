# 快速上手

前提：浏览器 DOM。本包的 import 对 SSR 安全（模块顶层不触碰任何 DOM 全局——仅渲染调用需要 `document`），但 `createWordPreview` 本身必须在浏览器里调用。

## 1. 安装

```bash
pnpm add @marcusok/word-preview
```

无需其他接线：渲染引擎（docx-preview）与其 zip 依赖（jszip）在构建期打进本包 `dist`。没有 WASM、没有 worker 脚本、不需要打包器插件。

## 2. 挂载预览

```ts
import { createWordPreview } from "@marcusok/word-preview";

const preview = createWordPreview(container, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onRendered: ({ pageCount, durationMs }) => {
    console.log(`渲染 ${pageCount} 个 section，耗时 ${durationMs}ms`);
  },
  onError: (e) => {
    // e.code: UNSUPPORTED | LEGACY_FORMAT | PASSWORD_PROTECTED | CORRUPT | UNKNOWN
    console.error(e.code, e.message);
  },
});
```

容器只是挂载点——预览渲染进内部 wrapper，因此 `destroy()` 即使在渲染仍在途时也能同步摘除。

## 3. 卸载

```ts
preview.destroy();
```

`destroy()` 立即生效：渲染 DOM 被摘除，之后才 resolve 的渲染结果被丢弃，所有回调静默。重复调用安全。同一容器可立即被新实例复用。

## 默认值与引擎的三处差异

三个默认值有意偏离 docx-preview 自身默认，按预览场景取舍：

| 选项                          | 引擎默认 | 本包                           | 原因                                                                                                         |
| ----------------------------- | -------- | ------------------------------ | ------------------------------------------------------------------------------------------------------------ |
| `useBase64URL`                | `false`  | `true`（锁定）                 | 引擎从不 revoke blob URL（上游 0 处 `revokeObjectURL`）；base64 规避泄漏。代价：DOM 内图片/字体内存约 +33%。 |
| `ignoreLastRenderedPageBreak` | `true`   | `false`（可覆盖）              | 更贴近 Word 存盘的分页。该标记由编辑器写入、非排版计算——与真实分页可能有出入。                               |
| `renderAltChunks`             | `true`   | `false`（见 `allowAltChunks`） | 引擎把 altChunk 渲染为无 sandbox 的 `srcdoc` iframe——对不可信文档是脚本执行向量。                            |

其余可透传项都在 `renderOptions` 里（见 [API](/zh/packages/word-preview/api/01-create-word-preview#wordrenderoptions)）。

## 下一步

- [限制](./02-limits)——主线程渲染、CSP 要求、CSS 注入、多实例残留。
- 在 [play](/zh/play) 里用自己的文件试用。
