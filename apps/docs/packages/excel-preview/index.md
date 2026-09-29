# @marcusok/excel-preview

Read-only xlsx preview in the browser, built on the same WASM engine as [@marcusok/excel-exporter](/packages/excel-exporter/): parsing runs in a Web Worker, rendering is a framework-agnostic DOM grid with virtual scrolling, and a dedicated compensation layer restores Excel-accurate number formatting and theme colors.

## Capabilities

| Capability             | Description                                                                                                                                                                       |
| ---------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Formats                | `.xlsx` / `.xlsm` (ZIP/OOXML), password-protected workbooks (`password` option), and `.csv` (UTF-8 / GB18030, delimiter sniffing)                                                 |
| Worker parsing         | The whole parse + model build runs in a shared Web Worker (~1.5s for 100k × 10 cells, measured; UI never freezes)                                                                 |
| Virtual scrolling      | DOM renderer mounts only viewport cells (plus a buffer); merged cells spanning the viewport are recalled as a whole                                                               |
| Layout fidelity        | Column widths, row heights, hidden rows/columns, merged cells, frozen panes (four-quadrant layers), sheet tabs (hidden sheets kept out)                                           |
| Style fidelity         | Fonts, solid / pattern / gradient fills, borders, alignment (wrap / indent / rotation), grid-line switch from the file                                                            |
| Theme & indexed colors | Recovered by a self-parsing overlay over `styles.xml` + `theme1.xml` — the engine's read path drops `theme`/`tint`/`indexed` colors                                               |
| Number-format fidelity | Excel's actual builtin table (not the ECMA strings), negative-sign / accounting-parens / currency-literal compensation, elapsed `[h]:mm:ss`, minute-adjacent `mm`, date1904 shift |
| Errors                 | Friendly, code-carrying errors: `PASSWORD_PROTECTED`, `LEGACY_FORMAT` (legacy .xls), `CORRUPT`, `UNSUPPORTED`, `WASM`, `UNKNOWN`                                                  |

## Install

```bash
pnpm add @marcusok/excel-preview
```

One same-scope dependency (`@marcusok/xlsx-core`, the shared engine layer): pages using both @marcusok spreadsheet packages load one engine instance and one WASM binary on the main thread. See [Quick Start](/packages/excel-preview/guide/01-quick-start).

## Quick example

```ts
import { createPreview } from "@marcusok/excel-preview";

const preview = createPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onParsed: (info) => console.log(info.sheetNames, info.duration),
  onError: (e) => console.error(e.code, e.message),
});

preview.setSheet("Sheet2"); // by name or index
preview.destroy(); // unmount + free resources
```

A low-level API — `parseWorkbookBytes(bytes, { password })` — returns a plain, structured-clone-safe data model (`PreviewWorkbook`) for React/Vue wrappers, SSR, or custom renderers.

<ClientOnly>
<PreviewDemo />
</ClientOnly>

## Not in scope (v1)

- Editing / formula recalculation — formula cells render their **cached** values (the industry convention; a formula saved without a cached value renders empty).
- Legacy `.xls` (BIFF8) and `.ods` — rejected with a clear message.
- Charts, images, shapes, conditional formatting, rich-text runs (rendered as plain concatenated text), and external hyperlink URLs (the engine's read path drops them; the link styling still shows).
