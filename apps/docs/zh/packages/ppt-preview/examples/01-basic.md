# 基础示例

最小而完整的接线：选文件、预览、导航、清理。

```ts
import { createPptPreview } from "@marcusok/ppt-preview";
import type {
  PptPreviewInstance,
  PptPreviewError,
} from "@marcusok/ppt-preview";

const container = document.querySelector<HTMLDivElement>("#preview")!;
const fileInput = document.querySelector<HTMLInputElement>("#file")!;
const status = document.querySelector<HTMLElement>("#status")!;

let preview: PptPreviewInstance | null = null;

fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (!file) return;

  // 挂载新实例前销毁旧实例。
  preview?.destroy();
  preview = null;
  container.textContent = "";

  preview = createPptPreview(container, {
    source: file,
    slideWidth: 960,
    onRendered: ({ slideCount, firstSlideMs, durationMs }) => {
      status.textContent = `共 ${slideCount} 页 · 首张 ${firstSlideMs}ms · 端到端 ${durationMs}ms`;
    },
    onError: (e: PptPreviewError) => {
      status.textContent = `${e.code}: ${e.message}`;
    },
  });
});

// 0-based 导航接到按钮上：
document
  .querySelector("#prev")
  ?.addEventListener("click", () => preview?.prevSlide());
document
  .querySelector("#next")
  ?.addEventListener("click", () => preview?.nextSlide());
document
  .querySelector("#first")
  ?.addEventListener("click", () => preview?.goToSlide(0));
```

标记 —— 容器必须真的会滚动（确定高度 + `overflow: auto`）：

```html
<input id="file" type="file" accept=".pptx" />
<button id="first">第一页</button>
<button id="prev">上一页</button>
<button id="next">下一页</button>
<span id="status"></span>
<div id="preview" style="height: 600px; overflow: auto"></div>
```

## 加密文件变体

```ts
preview = createPptPreview(container, {
  source: encryptedFile,
  password: passwordInput.value, // MS-OFFCRYPTO Agile「打开密码」
  onRendered,
  onError: (e) => {
    // e.code === "PASSWORD_PROTECTED"；e.message 区分
    //「请提供密码」与「密码被拒绝」。
  },
});
```

## 框架卸载清理

卸载时 destroy（React 示例）：

```tsx
useEffect(() => {
  if (!file || !ref.current) return;
  const p = createPptPreview(ref.current, { source: file });
  return () => p.destroy();
}, [file]);
```
