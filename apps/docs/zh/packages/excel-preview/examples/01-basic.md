# 基础用法

## 文件选择框

```ts
import { createPreview } from "@marcusok/excel-preview";

const input = document.querySelector<HTMLInputElement>("#file")!;
input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;
  createPreview(document.querySelector("#preview")!, {
    source: file,
    onParsed: (info) =>
      console.log(
        `${info.sheetCount} 个工作表，${info.rowCount} 行，解析 ${info.duration.parse}ms`,
      ),
    onError: (e) => alert(`${e.code}: ${e.message}`),
  });
});
```

## fetch 字节

```ts
const res = await fetch("/reports/2026-q3.xlsx");
const bytes = new Uint8Array(await res.arrayBuffer());
const preview = createPreview(el, { source: bytes });
```

## 加密工作簿

```ts
const preview = createPreview(el, {
  source: file,
  password: prompt("密码？") ?? undefined,
});
```

密码错误（或缺失）以 `onError` 的 `PASSWORD_PROTECTED` 错误码暴露——`createPreview` 本身从不抛异常。

## CSV

没有独立 API：非 ZIP 的文本形态输入自动按 CSV 解析（UTF-8 BOM 处理、中文导出的 GB18030 回退、`,` / `;` / tab / `|` 分隔符嗅探、RFC 4180 引号转义）。结果是单 sheet 的 `PreviewWorkbook`，符合 General 规则的数字单元格带 number 类型（右对齐语义）。

## 在线演示

<ClientOnly>
<PreviewDemo />
</ClientOnly>
