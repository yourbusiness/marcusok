# Basic Usage

## From a file input

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
        `${info.sheetCount} sheets, ${info.rowCount} rows, parsed in ${info.duration.parse}ms`,
      ),
    onError: (e) => alert(`${e.code}: ${e.message}`),
  });
});
```

## From fetched bytes

```ts
const res = await fetch("/reports/2026-q3.xlsx");
const bytes = new Uint8Array(await res.arrayBuffer());
const preview = createPreview(el, { source: bytes });
```

## Encrypted workbooks

```ts
const preview = createPreview(el, {
  source: file,
  password: prompt("Password?") ?? undefined,
});
```

A wrong (or missing) password surfaces as `onError` with code `PASSWORD_PROTECTED` — never a thrown exception from `createPreview` itself.

## CSV

No separate API: a non-ZIP, text-shaped input parses as CSV automatically (UTF-8 BOM handling, GB18030 fallback for Chinese exports, `,` / `;` / tab / `|` delimiter sniffing, RFC 4180 quoting). The result is a single-sheet `PreviewWorkbook` with number-typed cells where Excel's General rules would show them right-aligned.

## Live demo

<ClientOnly>
<PreviewDemo />
</ClientOnly>
