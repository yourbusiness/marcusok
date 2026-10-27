# 基础用法

## 上传 → 预览（React）

```tsx
import { useEffect, useRef, useState } from "react";
import {
  createPdfPreview,
  type PdfPreviewInstance,
} from "@marcusok/pdf-preview";

export function PdfPanel({ file }: { file: File }) {
  const ref = useRef<HTMLDivElement>(null);
  const previewRef = useRef<PdfPreviewInstance | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!ref.current) return;
    const preview = createPdfPreview(ref.current, {
      source: file,
      scale: "fit-width",
      onError: (e) => setError(`${e.code}: ${e.message}`),
    });
    previewRef.current = preview;
    return () => preview.destroy(); // 取消在途任务 + 终止 worker
  }, [file]);

  return (
    <div style={{ height: 600, overflow: "auto", background: "#f5f5f5" }}>
      {error && <p>{error}</p>}
      <div ref={ref} />
    </div>
  );
}
```

滚动容器是你自己的元素——预览在其中自然撑高，视口驱动的懒渲染自动跟随。观察器以最近的、真正会滚动或高度受约束的祖先为根（定高、`max-height` 或 flex 解析出的高度），无合适祖先时回退浏览器视口；仅仅随内容长高的 `overflow: auto` 包装层会被跳过，因此预载/释放带跟随的是实际滚动的元素。

## 缩放模式

```ts
createPdfPreview(el, { source, scale: "fit-width" }); // 页宽 = 容器宽（默认）
createPdfPreview(el, { source, scale: "fit-page" }); // 整页落入容器
createPdfPreview(el, { source, scale: 1.25 }); // pdf.js 原始 viewport 缩放
```

`fit-*` 创建时解析一次。之后改缩放就 `destroy()` 重建——首版有意不提供热重排 API。

## 文本层开关

```ts
createPdfPreview(el, { source, textLayer: false }); // 纯 canvas（不可选中）
```

文本层默认开启：透明 span 与 canvas 字形精确对齐，选中与复制像真实文本一样工作。

## 性能观测

```ts
createPdfPreview(el, {
  source,
  onRendered: ({ pageCount, firstPageMs, durationMs }) => {
    // pageCount：文档总页数；firstPageMs：首个绘制完成的页（canvas + 文本层）；
    // durationMs：从 createPdfPreview 调用到该页绘制完成
  },
});
```

200 页合成文本文档的解析 + 首页文本提取本地实测约 66ms（热态中位；一次性的 pdf.js worker/模块初始化由预热运行排除，见包 README）——终端用户真正感知的是 `firstPageMs`（含 canvas 绘制）。
