# 基础用法

## 渲染上传的文件

```ts
import { createWordPreview } from "@marcusok/word-preview";

const input = document.querySelector<HTMLInputElement>("#file")!;

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;

  const preview = createWordPreview(document.querySelector("#preview")!, {
    source: file,
    onRendered: ({ pageCount, durationMs }) =>
      console.log(`${pageCount} 个 section，完整渲染 ${durationMs}ms`),
    onError: (e) => {
      switch (e.code) {
        case "LEGACY_FORMAT":
          showHint(".doc 老格式与加密 docx 均不支持——请另存为 .docx");
          break;
        case "UNSUPPORTED":
          showHint("这不是 Word 文档");
          break;
        default:
          showHint(e.message);
      }
    },
  });

  // 之后：preview.destroy()
});
```

## 从字节渲染并传渲染选项

```ts
import { createWordPreview } from "@marcusok/word-preview";

const bytes: Uint8Array = await response
  .arrayBuffer()
  .then((b) => new Uint8Array(b));

const preview = createWordPreview(el, {
  source: bytes,
  renderOptions: {
    // 白名单引擎选项——共 13 项，见 API 页
    ignoreWidth: true,
    renderComments: false,
    // 把本包默认改回引擎自身行为
    ignoreLastRenderedPageBreak: true,
  },
});
```

## 同页两份文档

实例自动隔离各自的样式——直接创建两个即可：

```ts
const a = createWordPreview(leftPane, { source: fileA });
const b = createWordPreview(rightPane, { source: fileB });
// 各自渲染在自己的 docx-<n> 类名下；编号/主题 CSS 互不跨越
```

## React 式卸载

`destroy()` 同步且幂等；仍在途的渲染会写入已脱离的节点、结果被丢弃：

```tsx
useEffect(() => {
  const preview = createWordPreview(ref.current, { source });
  return () => preview.destroy();
}, [source]);
```
