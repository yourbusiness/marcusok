# Limits & Boundaries

Everything this package deliberately does **not** do, and every environment constraint. Read this before filing an issue — most "bugs" are documented trade-offs here.

## Format support

- **Legacy binary `.ppt` is not supported.** The engine confirms the legacy container and this package reports `LEGACY_FORMAT`. Convert to `.pptx` upstream.
- **Password decryption covers OLE2/MS-OFFCRYPTO Agile only** (PowerPoint's "password to open", the modern default). ZIP-entry encryption is not covered and is not promised.
- Decryption details you must know:
  - SHA-512 key derivation and AES-CBC run **on the main thread** via WebCrypto — a secure context (HTTPS / localhost) is required, and it typically adds **~1s** before parsing (spinCount-driven).
  - The upstream decryptor does **not verify the HMAC data-integrity tag**; tampered ciphertext is not detected by that mechanism.

## Loading semantics — lazy rendering ≠ lazy parsing

The engine's non-progressive load **preflights every slide** before `load()` resolves. Lazy rendering in this package only defers Canvas paint and the text layer; first-paint latency still grows with the total slide count. `onRendered.durationMs` therefore includes the full-deck preflight (this is also stated in the callback's doc comment).

**Transient memory during load:** input bytes are defensively copied twice — one copy is kept for sniffing, and a second dedicated buffer is handed to the parser (which transfers it into its worker). Peak source-related memory during loading is therefore ~2× the file size, transient; the caller's original bytes are never touched, and the copies become collectible once loading finishes.

## Media, hyperlinks, interactivity

- **Static preview only**: animations, slide transitions, and audio/video playback are not rendered. Media elements keep their static poster/graphics representation; media playback controls are explicitly skipped (`skipMediaControls`), so posters never show controls that look operational but are not.
- **Hyperlinks are inert** in v1: link text renders (and is selectable), but clicks do nothing. Opening links safely requires an external-scheme allowlist and internal slide-jump design (tracked as a follow-up).
- **Hidden slides are all shown**, counted in `slideCount`, and included in navigation — indices preserve the original deck order. No show/skip/dim configuration.

## Content degradation (optional engine renderers not injected)

This package intentionally does not inject the engine's optional renderer modules. Affected content degrades as follows:

| Content              | Degradation without the renderer                                                                                                                                                                                   |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| OMML equations       | Fallback representation (MathJax not loaded — no network font fetch)                                                                                                                                               |
| TIFF images          | Render error for recognized TIFF parts (browser decoder not used)                                                                                                                                                  |
| ChartEx charts       | Standard unsupported-chart placeholder                                                                                                                                                                             |
| 3-D chart groups     | Fall back to their canonical 2-D family                                                                                                                                                                            |
| Region maps          | Standard unsupported-chart placeholder                                                                                                                                                                             |
| SmartArt             | The PowerPoint-saved `dsp:drawing` is rendered when present; if missing, the engine falls back to a plain text list — there is no native diagram layout engine, so original SmartArt layouts are not reconstructed |
| Charts (classic 2-D) | Supported out of the box (included in the format entry)                                                                                                                                                            |

## Fonts & CSP

- **Embedded fonts** are honored and released on destroy. Google Fonts is **off by default** (`useGoogleFonts: false` — no implicit external requests); enabling it may fetch from `fonts.googleapis.com` (CSP: allow the domain; privacy: external requests).
- **CJK fallback**: `renderOptions.cjkFallback` (default `"auto"`) selects substitute glyphs for missing CJK fonts (`sc`/`tc`/`hk`/`jp`/`kr`).
- CSP requirements:
  - `worker-src blob:` and `worker-src data:` — the engine inlines its parser worker via a Blob URL (data: fallback). Opaque-origin iframes (sandboxed without `allow-scripts`-compatible origin) may break worker creation.
  - `script-src 'wasm-unsafe-eval'` (or equivalent) for WASM execution.
  - `style-src` (with the `styleNonce` you pass) — covers **this package's** injected layout/text-layer stylesheet.
  - The engine positions text-layer spans and this package sizes slide placeholders through CSSOM (`element.style` / `style.cssText`). In Chromium those writes are **not** blocked by `style-src-attr 'none'`, so `style-src-attr 'unsafe-inline'` is not required there. Firefox/WebKit were not verified; if your target blocks these writes, allow `style-src-attr 'unsafe-inline'` or set `enableTextSelection: false` (canvas only).

## Workers & multi-instance memory

- Each instance owns its parser Worker and a WASM instantiation. Multiple instances on one page multiply that cost — for example a main preview plus thumbnails. Main-thread rendering is serialized per instance, but parsing workers run concurrently.
- The two build-time warnings consumer bundlers may print about `pptx-source-worker.mjs` / `assets/render-worker-*.js` reference engine assets for the `mode: "worker"` rendering path and the legacy path — **this package never loads them** (`mode: "main"` fixed). The warnings are safe to ignore; suppressing them would require shipping ~2 MB of dead assets.
- **Vite 5–7 dev server**: dependency pre-bundling moves the package into `.vite/deps/`, so the wasm URL resolves to a missing file (HTML fallback/404) and loading reports `UNKNOWN`. Add `optimizeDeps: { exclude: ["@marcusok/ppt-preview"] }`; production builds are unaffected. See the quick start.

## Resource budgets

Valid files that exceed the engine's archive budgets (entry size, total inflated bytes, entry count, decoded image budgets) report `UNSUPPORTED` — they are **not** corrupt. Tune via `renderOptions.resourceLimits` (`maxArchiveEntryBytes`, `maxTotalInflatedBytes`, `maxArchiveEntries`); see the [API](/packages/ppt-preview/api/01-create-ppt-preview) page.

## Layout & sizing

- `slideWidth` is fixed at creation; there is **no resize re-layout** — destroy and re-create.
- `dpr` has no upper clamp: a large value multiplies the canvas bitmap quadratically (e.g. `dpr: 10` at the default 960-px width allocates ~9600×5400 bitmaps, ≈200 MB RGBA per slide). Only override it when you know the target displays.
- Empty decks (`slideCount = 0`) load without error and never fire `onRendered` — but the engine instance (parser Worker + WASM) stays alive until `destroy()`. Destroy empty-deck previews you are done with instead of leaving them mounted.
- SSR: the module can be imported on the server; `createPptPreview` must only run where `Worker` and WebCrypto exist.
- **Containers must live in the main document** — the wrapper creates slide/canvas/text-layer nodes through the global `document` and injects its layout/text-layer stylesheet into `document.head`. Rendering into a ShadowRoot or an iframe document is not supported in this version (the injected CSS would not reach the nodes, breaking placeholder layout and text selection).

## Error-normalization boundaries

- `not-ooxml` from the engine (non-pptx ZIPs, unreadable containers, some damaged encrypted containers) maps to `UNSUPPORTED` — the message is kept but damage causes are **not** further classified (the upstream error has no `stage` field to key on).
- Unknown parser traps, Worker/WASM/CSP and asset-URL failures map to `UNKNOWN` — never reported as file corruption.
- **A slide that fails to render reports an error but stays blank** — the slide-level failure path releases that slide's canvas and returns it to the empty state; it re-renders only when it leaves and re-enters the render band (IntersectionObserver only fires on crossing changes). The error message is the engine's and does not carry the slide index, so a per-slide failure is not distinguishable from a deck-level one by the error payload alone.
