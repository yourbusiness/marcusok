# @marcusok/ppt-preview

Read-only `.pptx` preview in the browser — continuous scrolling with lazy per-slide Canvas rendering, a selectable/copyable text layer, 0-based slide navigation, and password decryption for encrypted files. Rendering is powered by a Rust/WASM parser ([@silurus/ooxml](https://github.com/yukiyokotani/office-open-xml-viewer) 0.89.6, bundled into this package's `dist` — zero transitive dependencies for consumers), wrapped with format sniffing, defensive source copying and viewport-driven memory management.

## Capabilities

| Capability        | Description                                                                                                                                                                                                                                                           |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Formats           | Standard OOXML `.pptx`. Password-protected files (OLE2/MS-OFFCRYPTO **Agile** encryption) decrypt when `password` is provided. Legacy binary `.ppt` is detected and reported as `LEGACY_FORMAT` (never silently mis-rendered)                                         |
| Lazy rendering    | Continuous scroll; every slide is a placeholder sized by the deck's aspect ratio (no scrollbar jumps). Slides render as they approach the viewport (1-screen preload band) and release their canvas memory when far away (beyond 3 screens)                           |
| Text layer        | Text runs are collected **in the same render pass** (`onTextRun` — no double rendering); a transparent layer aligns with the canvas pixels — select and copy like real text. Off with `enableTextSelection: false`                                                    |
| Navigation        | `goToSlide(index)` / `nextSlide()` / `prevSlide()`, 0-based, tracking the most visible slide (manual scrolling is respected). Hidden slides are all shown and counted, preserving the original deck order                                                             |
| Worker parsing    | Parsing runs in a module Web Worker (WASM); slide layout and Canvas paint run on the main thread (`mode: "main"` upstream — the engine's fixed main-mode shape)                                                                                                       |
| Decryption        | MS-OFFCRYPTO Agile (SHA-512 key derivation + AES-CBC) via WebCrypto on the main thread — requires a secure context (HTTPS/localhost). Typically adds ~1s before parsing; ZIP-entry encryption is not covered                                                          |
| Hardened defaults | Sniffing runs **before** the engine (empty / non-ZIP / non-OLE2 inputs report early; OLE2 is passed through so the engine distinguishes legacy/encrypted/unknown CFB). Hyperlinks are inert (select/copy only); media renders as static posters (`skipMediaControls`) |
| Cancellation      | `destroy()` disconnects observers, expires in-flight renders, zeroes canvases, removes injected DOM and destroys the engine (terminating the parser worker); all callbacks go silent. Cancellation is not an error (no `onError`)                                     |
| Errors            | Code-carrying errors: `UNSUPPORTED`, `LEGACY_FORMAT`, `PASSWORD_PROTECTED`, `CORRUPT`, `UNKNOWN`                                                                                                                                                                      |

## Install

```bash
pnpm add @marcusok/ppt-preview
```

The engine (~2.5 MB JS) and its parser wasm (~2 MB, ~800 KB gzipped) ship inside this package's `dist`. Under Vite/webpack the wasm needs **zero configuration** — the bundler detects the asset next to the entry and emits it automatically. Two build-time warnings about worker assets (`pptx-source-worker.mjs`, `assets/render-worker-*.js`) are expected and safe to ignore: they belong to the engine's `mode: "worker"` rendering path and the legacy path, neither of which this package ever loads. See [Limits](/packages/ppt-preview/guide/02-limits).

## Quick example

```ts
import { createPptPreview } from "@marcusok/ppt-preview";

const preview = createPptPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  slideWidth: 960, // CSS px (default); height follows the deck ratio
  password: "…", // optional; MS-OFFCRYPTO Agile files only
  onRendered: ({ slideCount, firstSlideMs, durationMs }) =>
    console.log(slideCount, firstSlideMs, durationMs),
  onError: (e) => console.error(e.code, e.message),
});

preview.nextSlide(); // 0-based navigation
preview.destroy(); // cancel in-flight work + terminate the engine
```

There is intentionally no low-level data-model API: the engine's slide models are tied to a live parser session and cannot be honestly detached into a plain-data `parse()` result. This differs from [@marcusok/excel-preview](/packages/excel-preview/), which exposes `parseWorkbookBytes` — see its docs for that pattern.

<ClientOnly>
<PptPreviewDemo />
</ClientOnly>

## Not in scope (v1)

- **Legacy `.ppt`** — detected and reported as `LEGACY_FORMAT`; convert to `.pptx` upstream.
- **Animations, transitions, audio/video playback** — static preview only; media keeps its poster/graphics representation (no fake playback controls).
- **Hyperlinks** — rendered as text but inert (no click handling); opening the external-scheme allowlist and internal slide jumps is a follow-up.
- **Container re-layout** — `slideWidth` is resolved once at creation; container resizes do not re-fit (destroy + re-create instead).
- **Optional engine renderers** — MathJax (OMML equations), TIFF images, ChartEx, 3-D charts and Region Maps are not injected; affected content falls back to the engine's degradation (placeholder / 2-D family). See [Limits](./guide/02-limits).
