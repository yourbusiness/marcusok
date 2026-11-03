# @marcusok/pdf-preview

## 0.1.6

### Patch Changes

- 8c7eca9: Call the pdf.js `TextLayer.cleanup()` static with `TextLayer` as the receiver. The method reads `this.#pendingTextLayers`; the previous detached invocation threw `TypeError` on every last-instance destroy and was swallowed by the catch, so the global text-layer measurement cache was never actually cleaned.

## 0.1.5

### Patch Changes

- 1e3faca: Cap high-DPR canvas bitmaps at the pdf.js viewer's pixel budget (5,242,880) to avoid multi-hundred-MB page bitmaps; re-probe the scroll root after all placeholders are inserted and ignore percentage `max-height` wrappers; re-queue pages whose stale release event arrives after they re-entered the viewport; fix `fit-page` padding for content-box containers and warn when the height is unmeasurable; clean pdf.js text-layer caches when the last instance is destroyed; sniff before copying typed sources; and center oversized pages with `safe center` so their left edge stays reachable.

## 0.1.4

### Patch Changes

- c2151f5: Align the text layer on rotated pages (`/Rotate` 90/180/270 transform rules), make `fit-page` measure `max-height` containers before placeholders exist and reserve the wrapper's vertical padding so the page no longer overflows, report a missing `IntersectionObserver` as `UNKNOWN` (environment) instead of `CORRUPT`, and fall back to the packaged asset directories when `workerSrc` is a non-hierarchical `blob:`/`data:` URL.

## 0.1.3

### Patch Changes

- 9dfc708: Measure `fit-width`/`fit-page` against the container content box (padding excluded), apply PDF `/UserUnit` to the text layer, accept cross-realm byte sources and thenable callback rejections, and document the ShadowRoot/iframe and Vite 5–7 dev-server limits.

## 0.1.2

### Patch Changes

- 71229de: Detect scroll containers after page placeholders are mounted and isolate rejected promises from user callbacks.

## 0.1.1

### Patch Changes

- 272fed1: Fix lazy-render re-entry races that could leave a page permanently blank after scrolling away during an in-flight render and quickly scrolling back: visibility changes are now tracked while a page is rendering, cancelled PDF renders are re-queued when the page is visible again, and completed PPT renders are no longer released by a stale release request. Also clear stale PDF release state after failed renders, silence late callbacks after document-level failures, release the PPTX engine reference on load-time destroy, and normalize synchronous `load()` failures into `onError` instead of an unhandled rejection.

## 0.1.0

### Minor Changes

- 018e268: Initial release: read-only PDF preview in the browser, powered by pdf.js (pdfjs-dist 6.4.299 bundled — zero transitive dependencies).

  - `createPdfPreview(container, options)` one-line mounting with continuous scroll and viewport-driven lazy rendering (pages render as they approach the viewport, canvases release beyond 3 screens, renders serialized per instance; per-page placeholders sized from each page's own viewport so mixed-size documents don't jump)
  - Selectable/copyable text layer (pdf.js `TextLayer`) with the required positioning CSS injected as an idempotent singleton (`styleNonce` supported); `textLayer: false` for pure canvas
  - `scale`: number / `"fit-width"` (default) / `"fit-page"`, resolved once at creation
  - Worker + cmaps + standard_fonts ship in `dist`; under Vite/webpack the worker needs zero config (bundler-emitted anchor), cmaps/fonts are URL-overridable for bundler scenarios (CJK documents degrade silently without them)
  - Format sniffing before the engine (`%PDF-` magic within the first 1024 bytes, Adobe's reader window) and normalized error codes: `UNSUPPORTED` / `PASSWORD_PROTECTED` (detected, not decrypted) / `CORRUPT` / `UNKNOWN`
  - Hardened defaults: the pdf.js sandbox module is never loaded (embedded PDF JavaScript never executes), the annotation layer is not rendered, XFA is off
  - `destroy()` full-chain cancellation: observers disconnected, in-flight render tasks and text-layer streams cancelled, rendered canvases zeroed, worker terminated, DOM detached — destroy-path rejections are silent
  - Worker / asset load failures report `UNKNOWN` (with a worker-specific message) instead of masquerading as a corrupt PDF; explicit `workerSrc` may be absolute or relative (resolved against the page base before deriving cmaps/fonts)
