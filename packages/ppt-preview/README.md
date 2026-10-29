# @marcusok/ppt-preview

Read-only `.pptx` preview in the browser: continuous scrolling with lazy per-slide Canvas rendering, a selectable/copyable text layer, 0-based slide navigation, and password decryption for encrypted files. Powered by a Rust/WASM OOXML engine ([@silurus/ooxml](https://github.com/yukiyokotani/office-open-xml-viewer), bundled at build time — zero transitive dependencies).

## Highlights

- **Lazy continuous scroll** — slides render as they approach the viewport (±1 screen) and release their canvas memory beyond ±3 screens; placeholders are sized by the deck's aspect ratio (no scrollbar jumps).
- **Text layer in one pass** — text runs are collected during the same render pass as the canvas paint (no double rendering); select and copy like real text. `enableTextSelection: false` for pure canvas.
- **0-based navigation** — `goToSlide` / `nextSlide` / `prevSlide` track the most visible slide; hidden slides are all shown and counted in the original deck order.
- **Encrypted files decryptable** — MS-OFFCRYPTO Agile ("password to open") via WebCrypto; missing vs. wrong passwords are distinguishable error codes. Main-thread decryption, ~+1s (spinCount-driven), secure context required.
- **Hardened defaults** — format sniffing before the engine, defensive copies of your input bytes (the parser transfers its buffer), inert hyperlinks, media as static posters, normalized error codes (`UNSUPPORTED` / `LEGACY_FORMAT` / `PASSWORD_PROTECTED` / `CORRUPT` / `UNKNOWN`).

## Install

```bash
pnpm add @marcusok/ppt-preview
```

ESM-only, Node >= 22. The engine (~2.5 MB JS) and parser wasm (~2 MB, ~800 KB gzipped) ship inside `dist`. Under Vite/webpack the wasm needs **zero configuration**; two build-time warnings about worker assets are expected and safe to ignore (they reference the engine's worker-mode/legacy paths this package never loads — see the docs' Limits page).

## Usage

```ts
import { createPptPreview } from "@marcusok/ppt-preview";

const preview = createPptPreview(container, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  slideWidth: 960, // CSS px, default; height follows the deck ratio
  onRendered: ({ slideCount, firstSlideMs, durationMs }) =>
    console.log(slideCount, firstSlideMs, durationMs),
  onError: (e) => console.error(e.code, e.message),
});

preview.nextSlide();
preview.destroy(); // idempotent; expires in-flight work + terminates the engine
```

## Performance (measured, 2026-10)

Measured locally in real Chromium on an Apple-silicon Mac; the deck load includes the engine's **full-deck preflight** (non-progressive loads parse every slide before resolving — lazy rendering defers painting, not parsing):

| Scenario                                  | First slide | End-to-end                                 |
| ----------------------------------------- | ----------- | ------------------------------------------ |
| Synthetic 30-slide deck (16:9, plain)     | ~21 ms      | ~104 ms                                    |
| Real 3-slide deck                         | ~5–12 ms    | ~83–92 ms                                  |
| Encrypted deck (spinCount 1000)           | ~5 ms       | ~67 ms                                     |
| Real Office encryption (spinCount 100000) | —           | ~+1.0–1.15 s of main-thread key derivation |

Node parsing baseline (regression gate in `performance.test.ts`): 200-slide full parse, median ~3 ms, threshold 50 ms.

## Limitations (v1)

Legacy `.ppt`, animations/transitions, media playback, clickable hyperlinks and container resize re-layout are out of scope; optional engine renderers (MathJax, TIFF, ChartEx, 3-D, Region Maps) are not injected and degrade gracefully. Full list on the [docs Limits page](https://yourbusiness.github.io/marcusok/packages/ppt-preview/guide/02-limits).

## Documentation

Docs site: https://yourbusiness.github.io/marcusok/ (en + zh).

## License

MIT. The bundled engine's MIT license and third-party notices are in [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md).
