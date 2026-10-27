# @marcusok/pdf-preview

Read-only PDF preview in the browser — continuous scrolling with lazy per-page rendering, a selectable/copyable text layer, and normalized error codes. Rendering is powered by [pdf.js](https://github.com/mozilla/pdf.js) (pdfjs-dist 6.4.299, bundled into this package's `dist` — zero transitive dependencies for consumers), wrapped with format sniffing, asset URL derivation and viewport-driven memory management.

## Capabilities

| Capability        | Description                                                                                                                                                                                                                                                                                         |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Formats           | Any spec-compliant PDF: the `%PDF-` magic is searched in the first 1024 bytes (Adobe's own reader window — junk bytes before the header are allowed). Files whose header sits beyond that window report `UNSUPPORTED` (a deliberate input tightening)                                               |
| Lazy rendering    | Continuous scroll; each page is a placeholder sized by its **own** viewport (no scrollbar jumps on mixed-size documents). Pages render as they approach the viewport (1-screen preload band) and release their canvas memory, text-layer stream and page resources when far away (beyond 3 screens) |
| Text layer        | A transparent text layer aligned with the canvas pixels (pdf.js `TextLayer`) — select, copy and search the preview like real text. Off with `textLayer: false`                                                                                                                                      |
| Scaling           | `scale`: a number, `"fit-width"` (default) or `"fit-page"`, resolved once at creation from the container and the first page; pages of differing sizes stay centered and proportional                                                                                                                |
| Worker parsing    | Parsing and rendering orchestration run in a pdf.js Web Worker; only canvas paint and DOM updates touch the main thread                                                                                                                                                                             |
| Hardened defaults | Sniffing runs **before** the engine; encrypted PDFs report `PASSWORD_PROTECTED` (no decryption in v1); the pdf.js sandbox module is never loaded, so embedded PDF JavaScript never executes; the annotation layer (links, form widgets) is not rendered in v1                                       |
| Cancellation      | `destroy()` disconnects observers, cancels in-flight render tasks and text-layer streams, zeroes rendered canvases, terminates the worker (`loadingTask.destroy()`); all callbacks go silent. Cancellation is not an error (no `onError`)                                                           |
| Errors            | Code-carrying errors: `UNSUPPORTED`, `PASSWORD_PROTECTED`, `CORRUPT`, `UNKNOWN`                                                                                                                                                                                                                     |

## Install

```bash
pnpm add @marcusok/pdf-preview
```

The pdf.js engine, its worker and the CJK/standard-font asset directories all ship inside this package's `dist` (~4.5 MB payload + sourcemap; `cmaps` are a hard requirement for Chinese/CJK PDFs). Under Vite/webpack the worker needs **zero configuration**; the `cmaps`/`standard_fonts` directories need either explicit URLs or a copy plugin in bundler scenarios — see [Quick Start](/packages/pdf-preview/guide/01-quick-start).

## Quick example

```ts
import { createPdfPreview } from "@marcusok/pdf-preview";

const preview = createPdfPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  scale: "fit-width", // default; number | "fit-page" also accepted
  onRendered: ({ pageCount, firstPageMs, durationMs }) =>
    console.log(pageCount, firstPageMs, durationMs),
  onError: (e) => console.error(e.code, e.message),
});

preview.destroy(); // cancel in-flight work + terminate the worker
```

There is intentionally no low-level data-model API: pdf.js exposes its document model through promise-based proxies tied to a live worker, which cannot be honestly detached into a plain-data `parse()` result. This differs from [@marcusok/excel-preview](/packages/excel-preview/), which exposes `parseWorkbookBytes` — see its docs for that pattern.

<ClientOnly>
<PdfPreviewDemo />
</ClientOnly>

## Not in scope (v1)

- **Decryption** — password-protected PDFs are detected and reported (`PASSWORD_PROTECTED`); no password prompt, no decryption.
- **Annotation layer** — hyperlinks, form widgets and comments do not render (attack surface and scope control; tracked as a follow-up).
- **Container re-layout** — the scale is resolved once at creation; container resizes do not re-fit (destroy + re-create instead).
- **JPX/JBIG2 images & ICC color management** — the pdf.js wasm assets are intentionally not bundled; affected documents degrade (JS fallback codecs / no color management) unless you self-host and pass `wasmUrl`/`iccUrl`. See [Limits](./guide/02-limits).
