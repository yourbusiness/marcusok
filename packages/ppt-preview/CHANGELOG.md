# @marcusok/ppt-preview

## 0.1.4

### Patch Changes

- 9a80c05: Copy `Buffer`/`Uint8Array` sources by offset and length (no memory-pool exposure), validate `cjkFallback` synchronously, guard lazy rendering after destroy, accept cross-realm sources and thenable callbacks, and correct the CSP, navigation, cause, hidden-container and `slideWidth` documentation.

## 0.1.3

### Patch Changes

- 71229de: Detect scroll containers after slide placeholders are mounted, narrow the source-only WASM fallback to the actual assets module, and isolate rejected promises from user callbacks.

## 0.1.2

### Patch Changes

- 272fed1: Fix lazy-render re-entry races that could leave a page permanently blank after scrolling away during an in-flight render and quickly scrolling back: visibility changes are now tracked while a page is rendering, cancelled PDF renders are re-queued when the page is visible again, and completed PPT renders are no longer released by a stale release request. Also clear stale PDF release state after failed renders, silence late callbacks after document-level failures, release the PPTX engine reference on load-time destroy, and normalize synchronous `load()` failures into `onError` instead of an unhandled rejection.

## 0.1.1

### Patch Changes

- 198fd99: Fix an unhandled rejection when post-load initialization fails (e.g. missing `IntersectionObserver` in very old WebViews, or SSR misuse): the error is now normalized and reported via `onError` with full cleanup instead of escaping as an unhandled promise rejection, matching the boot contract of the other preview packages.

## 0.1.0

### Minor Changes

- a8de23f: Add `@marcusok/ppt-preview` — read-only .pptx preview in the browser. Continuous scroll with lazy per-slide Canvas rendering (viewport-adjacent slides render, far-away slides release canvas memory), a selectable/copyable text layer collected in the same render pass, 0-based slide navigation (`goToSlide`/`nextSlide`/`prevSlide`) tracking the most visible slide, password decryption for MS-OFFCRYPTO Agile encrypted files, and normalized error codes. Powered by a Rust/WASM engine (@silurus/ooxml, bundled at build time — zero transitive dependencies; the parser wasm ships in `dist`).
