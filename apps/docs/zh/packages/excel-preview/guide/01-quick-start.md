# 快速上手

一个 `createPreview` 调用挂载完整预览：Worker 解析、错误归一与虚拟滚动渲染。

## 最小示例

```ts
import { createPreview } from "@marcusok/excel-preview";

const el = document.querySelector("#preview")!;
const preview = createPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
});

// 用完
preview.destroy();
```

这就是全部接入。浏览器内解析默认走 Web Worker；Node/SSR（无 Worker 全局）同一条管线自动在主线程跑——业务代码零改动。

## 常用选项

```ts
const preview = createPreview(el, {
  source: bytes,
  password: "…", // 加密工作簿（Agile AES-256）
  sheet: "Summary", // 初始 sheet：名称或 0 起索引
  showHeaders: false, // 隐藏 A/B/C + 1/2/3 表头
  showGridLines: false, // 覆盖文件声明的网格线开关
  showTabs: false, // 隐藏 sheet 页签栏
  onParsed: (info) => {
    // sheetNames / sheetCount / rowCount / colCount / duration{parse,render,total}
  },
  onError: (e) => {
    // e.code: PASSWORD_PROTECTED | LEGACY_FORMAT | CORRUPT | UNSUPPORTED | WASM | UNKNOWN
  },
});
```

## 切换 sheet

```ts
preview.setSheet(1); // 按索引
preview.setSheet("Sheet2"); // 按名称
preview.getSheetNames(); // 全部 sheet（含隐藏，按文件顺序）
```

默认渲染的页签栏在视觉上做同一件事；文件中标记 `hidden` / `veryHidden` 的表永远不会出现在页签里。

时机很重要：这两个方法只在**解析完成后**才生效。`createPreview` 是异步启动的，在它返回后的下一行就调用 `setSheet()` 属于静默空操作——没有返回值、也不触发 `onError`；在此期间 `getSheetNames()` 返回 `[]`。请在 `onParsed` 里、或由你自己的 UI、或交给页签栏来调用。`destroy()` 之后两者同样变回空操作；注意 `destroy()` **不会**中止在途解析，它只是抑制该次解析本应触发的 `onParsed` / `onError` 回调。

## 低层解析 API

不想用内置渲染器——自研 UI、SSR 或框架封装时：

```ts
import { parseWorkbookBytes } from "@marcusok/excel-preview";

const workbook = await parseWorkbookBytes(bytes, { password: "…" });
// workbook.sheets[0].rows[0].cells[0] → { col, type, value, styleIndex }
// styles: fonts / fills / borders / xfs（numFmtCode 已解析）
```

模型是纯 JSON（结构化克隆安全），与 Worker 传回渲染器的数据完全一致。字段细节见[模型类型](/zh/packages/excel-preview/api/02-model)。

## 下一步

- [资产与自托管](/zh/packages/excel-preview/guide/02-assets)——WASM/worker 文件如何定位、什么时候需要 `configureWasm`。
- [格式忠实度](/zh/packages/excel-preview/guide/03-format-fidelity)——补偿层还原了什么、哪些是已知近似。
- [能力边界](/zh/packages/excel-preview/guide/04-limits)——v1 明确不做的事。
