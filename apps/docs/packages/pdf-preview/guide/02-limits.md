# Limits & Security Notes

## Rendering

- **Lazy pipeline** — pages render when they enter a 1-screen preload band and their canvas is released when they leave a 3-screen band. Fast scrolling through a long document renders only the pages that settle; a page queued but no longer wanted is dropped before rendering.
- **One page at a time** — renders are serialized per instance (scroll bursts never fan out into parallel canvas work). Interleaved scrolling prioritizes whatever the observer reports next.
- **Placeholder accuracy** — every page's placeholder is sized from that page's own viewport, so mixed-size documents (e.g. US Letter + A4 landscape) never jump. Page proxies for all pages are fetched up-front so total scroll height is correct immediately after mount.
- **Canvas crispness** — canvases are painted at `devicePixelRatio` resolution with CSS-size placeholders, matching the official viewer.
- **Scale is resolved once** — `fit-*` measures the container at creation. `fit-page` needs a measurable, height-constrained container (fixed height, `max-height`, or a flex-resolved height); in a container that grows with its content the measured height is not a viewport, so give it a fixed height or use `fit-width`. Container resizes never re-fit.
- **Scrollbar timing** — `fit-width` measures the container **before** the page placeholders exist, i.e. before a vertical scrollbar appears in the scroll container. On platforms with classic (non-overlay) scrollbars the content box then narrows by the scrollbar width (~15–17 px) and the centered pages can overflow horizontally by that much. Re-creating the preview does not help (the measurement repeats in the empty state). Give the container `scrollbar-gutter: stable` or account for the gutter in your layout if this matters.
- **`devicePixelRatio` is captured per render** — moving the window across screens with a different DPR (or zooming) does not re-render existing pages; already-rendered pages keep their old bitmap sharpness until they are released and re-rendered.

## Formats

- The `%PDF-` magic must appear within the **first 1024 bytes** (Adobe's own reader window). Adobe's spec allows junk before the header and the official viewer searches a wider window — pathological files with >1 KB of junk open there but report `UNSUPPORTED` here. This is a deliberate input tightening.
- **Encrypted PDFs** are detected and reported as `PASSWORD_PROTECTED` (message distinguishes "needs a password" from "password rejected" upstream codes). No decryption in v1.
- **XFA forms** do not render (`enableXfa` is pinned to `false`); XFA-only documents show blank pages.
- **Node.js server-side calls are not supported**: the modern pdf.js build (which this package bundles) crashes in `getDocument` under Node. Importing the package is safe; calling `createPdfPreview` requires a browser. Server-side text extraction is out of scope.

## Worker fallback (silent degradation)

Parsing normally runs in a pdf.js Web Worker, but pdf.js itself silently falls back to a **main-thread fake worker** when the worker cannot be constructed — a CSP `worker-src` block, a wrong MIME type, or a module-worker restriction in the host. In that mode parsing still works (usually a `Setting up fake worker.` warning in the console at most) but happens on the main thread: long documents will jank the UI. Two things to know:

- pdf.js's "worker disabled" flag is **sticky for the page session** — once any construction fails, every later instance on the page (from this or any other pdf.js copy) uses the fake worker until reload.
- Only when even the main-thread import of the worker module fails does the package surface it as an `UNKNOWN` error mentioning `workerSrc`. If you require worker isolation, verify in DevTools that a `pdf.worker` thread exists rather than trusting the render to succeed.

## Security hardening (what the wrapper does)

- Sniffing runs **before** the engine: a non-PDF gets a readable `UNSUPPORTED` error instead of an opaque worker failure.
- **Embedded PDF JavaScript never executes**: pdf.js 6.x moved script execution into its sandbox module (QuickJS), which this package never loads. (The old `isEvalSupported` switch no longer exists upstream.)
- The **annotation layer** (hyperlinks, launch actions, form widgets) is not rendered in v1, removing the main interaction-based attack surface. Annotation _appearances_ (flattened stamps, watermarks) still paint onto the canvas — that is upstream's default render mode; what v1 skips is the interactive layer.
- Rendering output is canvas pixels + DOM text nodes — no iframes, no `srcdoc`, no HTML injection paths.

## wasm-gated features (not bundled)

The pdf.js wasm assets (`jbig2.wasm`, `openjpeg.wasm`, `qcms_bg.wasm`) and ICC profiles are intentionally **not** redistributed. Without `wasmUrl`:

| Content                          | Behavior without wasm                                   |
| -------------------------------- | ------------------------------------------------------- |
| JPEG2000 (JPX) images            | JS fallback decoder (slower) if available, else skipped |
| JBIG2 images (scanned documents) | JS fallback decoder (slower) if available, else skipped |
| ICC color-managed images         | colors render without color management (upstream warns) |

To enable full fidelity, self-host the `wasm/` (and optionally `iccs/`) directories from the `pdfjs-dist` npm package and pass `wasmUrl` / `iccUrl`. Missing wasm degrades **silently** — another reason to smoke-test with real documents.

## CSP requirements

- `font-src data:` — pdf.js injects embedded fonts as `data:font/...` URLs.
- `style-src` (sheet level) — pdf.js's FontLoader injects `@font-face` rules into `adoptedStyleSheets` or a `<head><style>` at runtime. This is upstream behavior and is **not** covered by the package's `styleNonce`.
- `style-src-attr 'unsafe-inline'` — text-layer spans and page placeholders use inline `style` attributes (the `styleNonce` option only covers the package's own injected stylesheet).
- `worker-src`/`script-src` — the pdf.js worker is a module worker loaded same-origin (or from wherever you point `workerSrc`); a cross-origin `workerSrc` makes pdf.js wrap the import in a blob worker.

## Memory

- Released pages cancel any in-flight text-layer stream, zero the canvas bitmap (`width = height = 0`), drop the text-layer DOM, and call pdf.js's `page.cleanup()` to release main-thread operator lists/objects. Re-entering the viewport re-renders; page data is fetched again on demand (the PDF is not re-parsed).
- `destroy()` additionally cancels in-flight render tasks and text streams, zeroes rendered canvases and clears the wrapper before detaching it.
- The document proxy and worker stay alive until `destroy()`; a preview that is removed from the DOM without `destroy()` leaks the worker. A document-level load failure (the `onError` path) tears the worker down immediately on its own — only a successfully mounted preview holds one.
- All-page proxy prefetch (needed for accurate placeholders) keeps one page proxy per page for the document's lifetime — that is pdf.js's normal `getPage` cache behavior. Canvas bitmaps and operator lists do not accumulate because release cleans them.

## SSR

Import-safe (module top level evaluates no DOM globals), call-unsafe: `createPdfPreview` touches `document`/`IntersectionObserver` immediately. Guard it in `onMounted`/`useEffect`, like any DOM library. Importing the package in Node does make pdf.js's modern build log a one-time `Please use the legacy build in Node.js environments` warning — expected, and it does not make server-side rendering work.

## Multi-instance

Instances are fully isolated in their containers. The injected stylesheet (page layout + text layer) is a document-level idempotent singleton shared by all instances and intentionally survives `destroy()`.

- pdf.js's `GlobalWorkerOptions.workerSrc` is module-level: if instances pass different `workerSrc` values, the last created instance wins. Use the same worker URL (or the default) across instances on one page.
- The stylesheet singleton is injected once; the first instance's `styleNonce` wins. Use the same nonce for all instances, or omit it.
