# @marcusok/pdf-preview

Read-only PDF preview in the browser — continuous scrolling with lazy per-page rendering, a selectable/copyable text layer and normalized error codes. Rendering is powered by [pdf.js](https://github.com/mozilla/pdf.js) (pdfjs-dist 6.4.299, Apache-2.0) bundled into this package's `dist`, together with the worker script and the CMap / standard-font asset directories — **zero transitive dependencies** for consumers.

```ts
import { createPdfPreview } from "@marcusok/pdf-preview";

const preview = createPdfPreview(container, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  scale: "fit-width", // number | "fit-page" also accepted
  onRendered: ({ pageCount, firstPageMs, durationMs }) =>
    console.log(pageCount, firstPageMs, durationMs),
  onError: (e) => console.error(e.code, e.message),
});

preview.destroy(); // cancel in-flight work + terminate the worker
```

## Why the wrapper (what you get over raw pdfjs-dist)

| Concern                 | Raw pdfjs-dist                                                                                    | This package                                                                                                                                                                                       |
| ----------------------- | ------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Bad input               | A non-PDF dies inside the worker with an opaque error                                             | Sniffed before the engine: `%PDF-` magic in the first 1024 bytes (Adobe's own reader window) → readable `UNSUPPORTED` / `CORRUPT` codes                                                            |
| Asset wiring            | `GlobalWorkerOptions.workerSrc` must be set by hand or the API throws; cmaps/fonts must be hosted | Worker located by a `new URL(..., import.meta.url)` anchor (bundler-emitted, zero config) + derived cmaps/fonts URLs, all overridable                                                              |
| Text layer styling      | `TextLayer` ships **no** positioning CSS — spans need the official viewer's `.textLayer` rules    | Minimal style set injected as a document-level idempotent singleton (with the `--scale-factor` variable chain), `styleNonce` supported                                                             |
| Large documents         | Every page you render stays rendered; canvas memory grows linearly                                | Viewport-driven lazy pipeline: 1-screen preload band renders, beyond 3 screens canvases are released (text streams cancelled, page resources cleaned) and re-render on return; serialized per page |
| Scrollbar jumps         | N/A (you build the layout)                                                                        | Per-page placeholders sized from each page's own viewport — mixed-size documents don't jump; devicePixelRatio-aware canvas paint                                                                   |
| Encrypted PDFs          | `PasswordException` surfaces raw (code 1/2 semantics undocumented at the call site)               | Normalized `PASSWORD_PROTECTED` with a message distinguishing needs-password from rejected-password                                                                                                |
| destroy()               | Nothing to call but `loadingTask.destroy()`; in-flight renders reject into your console           | Full-chain cancellation: observers disconnected, render tasks and text streams cancelled, canvases zeroed, worker destroyed, DOM detached — destroy-path rejections are silent                     |
| Embedded PDF JavaScript | Runs via the QuickJS sandbox if you load the sandbox module                                       | The sandbox module is never loaded — embedded PDF JavaScript never executes, unconditionally                                                                                                       |

## Limits (stated honestly)

- **Node/server-side calls are not supported** — the modern pdf.js build bundled here crashes in `getDocument` under Node (verified). Importing the package is SSR-safe; calling `createPdfPreview` requires a browser.
- **No decryption** — encrypted PDFs are detected and reported, never decrypted.
- **No annotation layer in v1** — hyperlinks, form widgets and comments do not render.
- **Scale is resolved once** at creation (`fit-*` measures the container then); container resizes don't re-fit — re-create the instance.
- **wasm-gated features degrade silently** — JPX/JBIG2 images and ICC color management need the pdf.js wasm assets, which are intentionally not bundled; pass `wasmUrl`/`iccUrl` if you need them.
- **Bundler scenarios and cmaps** — bundlers emit the worker automatically but never copy the `cmaps/`/`standard_fonts/` directories; without explicit URLs, CJK documents degrade silently. See the docs-site quick start for the copy-plugin recipe.

Full details in the [docs site limits page](https://yourbusiness.github.io/marcusok/packages/pdf-preview/guide/02-limits).

## Performance

- **Bundle** (measured at build time, unminified per repo convention): `dist/index.js` ~850 KB (~177 KB gzipped), worker 1.26 MB, `cmaps/` 1.6 MB (169 files), `standard_fonts/` 816 KB (16 files) — ~4.5 MB payload plus sourcemap, in the same class as `@marcusok/excel-preview`.
- **Node parse baseline** (`performance.test.ts`): 200-page synthetic text PDF, `getDocument` + first-page `streamTextContent` ≈ 66 ms median on the dev machine; SLA 150 ms (`PERF_SLACK` honored, `RUN_PERF=0` skips).
- **Rendering** is viewport-driven and serialized by design; `onRendered.firstPageMs` is the user-perceived first-paint number for your documents.

## Install

```bash
pnpm add @marcusok/pdf-preview
```

Requires Node `>=22.13.0 || >=24` at install time (the pdf.js engine's own declaration); the browser runtime only needs a modern evergreen browser.

## License

MIT. This package redistributes pdf.js (Apache-2.0, Mozilla Foundation), the Foxit/Liberation standard fonts and Adobe's CMap data — see [`THIRD_PARTY_NOTICES.md`](./THIRD_PARTY_NOTICES.md) and the `LICENSE*` files shipped inside the asset directories.
