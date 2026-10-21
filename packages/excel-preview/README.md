# @marcusok/excel-preview

Read-only xlsx preview in the browser, built on the same WASM engine as
[@marcusok/excel-exporter](https://www.npmjs.com/package/@marcusok/excel-exporter):
parsing runs in a Web Worker, rendering is a framework-agnostic DOM grid with
virtual scrolling, and a dedicated compensation layer restores Excel-accurate
number formatting and theme colors.

```ts
import { createPreview } from "@marcusok/excel-preview";

const preview = createPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onParsed: (info) => console.log(info.sheetNames, info.duration),
  onError: (e) => console.error(e.code, e.message),
});

// setSheet() / getSheetNames() only take effect once the parse resolves
preview.setSheet("Sheet2"); // by name or index
preview.destroy(); // also aborts an in-flight parse (ABORTED — never an onError)
```

That is the entire setup. Parsing runs in a Web Worker in the browser; the
same pipeline runs on the main thread with no code change whenever no worker
is usable — Node/SSR (no Worker **or** no `window` global), or a worker that
failed to construct/load (a permanent fallback for the session; see the assets
guide). `onParsed` fires after the first render
and again after every `setSheet`/tab switch (`duration.parse` reuses the first
parse's timing; `duration.total`, by contrast, keeps counting from the initial
load, so use `duration.render` for a switch's own cost).

## What it restores

- **Values & types** — numbers, strings, booleans, error codes, and formula
  cells' **cached** values (the industry convention; formulas are never
  recalculated).
- **Layout** — column widths, row heights, hidden rows/columns, merged cells
  (viewport-spanning merges are recalled whole), frozen panes (four-quadrant
  layers), sheet tabs.
- **Styles** — fonts, solid/pattern/gradient fills, borders, alignment
  (wrap/indent/rotation), grid-line switch; **theme (`theme`+`tint`) and
  indexed colors** are recovered by a self-parsing overlay over
  `styles.xml`/`theme1.xml`, because the engine's read path drops them.
- **Number formats** — builtin ids through Excel's actual behavior table
  (`m/d/yyyy`, not the ECMA `mm-dd-yy`), negative-sign / accounting-parens /
  currency-literal compensation, elapsed `[h]:mm:ss` durations,
  minute-adjacent `mm` (the engine renders it as month), `date1904` shift,
  `[Red]`-style section colors.
- **More** — password-protected workbooks (`password` option), CSV (UTF-8 /
  GB18030, delimiter sniffing), parse cancellation (`destroy()` aborts an
  in-flight parse; `parseWorkbookBytes` accepts an `AbortSignal`), friendly
  code-carrying errors (`PASSWORD_PROTECTED`, `LEGACY_FORMAT` for legacy
  `.xls`, `ABORTED` for cancellations, …).

## Low-level API

`parseWorkbookBytes(bytes, { password, signal })` returns the plain,
structured-clone-safe data model (`PreviewWorkbook`) with no DOM involved —
for framework wrappers, SSR, or custom renderers. Custom renderers can also
reuse `formatCellValue` (exported from the main entry), the same cell
formatter the DOM renderer uses.

## Large files

Parsing is not streaming (OOXML parts cross-reference each other), but it
happens entirely inside a Web Worker — the UI never freezes — and the DOM
renderer virtualizes to viewport-only cells. Measured locally: ~1.5s parse for
100k rows × 10 columns; scrolling cost is independent of file size.

## Not in scope

Editing, formula recalculation, legacy `.xls` (BIFF8) / `.ods`, charts,
images, shapes, conditional formatting, rich-text runs (rendered as plain
text), external hyperlink URLs (not rendered), split panes (only freeze panes
render), tab colors, and a pure-JS fallback
reader for environments without WebAssembly (the preview fails with a clear
`WASM` error instead; CSV still parses without WebAssembly). See the
docs site for the full
[limits list](https://yourbusiness.github.io/marcusok/packages/excel-preview/guide/04-limits).

## Assets

The WASM binary (1.9MB, gzip ≈ 650KB) and the self-contained parse worker
ship with this package and resolve automatically in bundlers and Node — the
engine is bundled into this package's `dist` at build time, so each @marcusok
spreadsheet package carries its own copy (identical binaries, usually
deduplicated as an asset by content-hash naming and the HTTP cache).
Self-hosted copies use the same `configureWasm({ wasmUrl, parseWorkerUrl })`
call shape as the exporter (the parse worker has its own `parseWorkerUrl`
field — each package's worker is a different script, so one field could not
serve both). The parse worker is a module-level singleton that reads
`parseWorkerUrl` once at creation: after changing URLs, call
`terminateParseWorker()` (exported from the main entry) — or reload the page —
before the next parse. In-flight parses are rejected without a main-thread
rerun; the next parse rebuilds a fresh worker from the current options.

## Documentation

- [Docs site](https://yourbusiness.github.io/marcusok/) (English + 中文)

## License

MIT
