# Quick Start

## Install

```bash
pnpm add @marcusok/ppt-preview
# npm i @marcusok/ppt-preview / yarn add @marcusok/ppt-preview both work
```

The package is ESM-only and bundles its rendering engine (Rust/WASM, `@silurus/ooxml` 0.89.6) at build time — consumers get **zero transitive dependencies**. Node ≥ 22.

> **Vite 5–7 dev server.** Production builds (Vite 8+, webpack 5) emit the parser wasm automatically, but Vite 5–7 pre-bundles npm-installed dependencies with esbuild into `node_modules/.vite/deps/`. The pre-bundled copy keeps `new URL("pptx_parser_bg.wasm", import.meta.url)` verbatim, so the wasm request resolves into `.vite/deps/` where the file does not exist (Vite answers with its HTML fallback or 404) and the preview reports `UNKNOWN`. Add the package to `optimizeDeps.exclude` and restart the dev server:
>
> ```ts
> // vite.config.ts
> export default { optimizeDeps: { exclude: ["@marcusok/ppt-preview"] } };
> ```
>
> Production builds are unaffected. (Vite 8's rolldown-based optimizer rewrites the literal correctly; linked workspace packages are not pre-bundled.)

## Render a presentation

```ts
import { createPptPreview } from "@marcusok/ppt-preview";

const container = document.querySelector<HTMLDivElement>("#preview");

const preview = createPptPreview(container!, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onRendered: ({ slideCount, firstSlideMs, durationMs }) => {
    console.log(
      `${slideCount} slides; first slide ${firstSlideMs}ms, end-to-end ${durationMs}ms`,
    );
  },
  onError: (e) => console.error(e.code, e.message), // see error codes below
});
```

The container should be a real scrolling element (`overflow: auto` with a definite height, e.g. a fixed-height panel). The preview is a vertical stack of slide-sized placeholders; slides render lazily as they enter the viewport ±1 screen and release their canvas memory beyond ±3 screens.

## Sizing

- `slideWidth` (default `960`) — the CSS pixel width of each slide; the height follows the deck's own aspect ratio (16:9, 4:3, …). Invalid values throw `TypeError` synchronously.
- `dpr` — override the device pixel ratio (defaults to `window.devicePixelRatio`). The canvas bitmap is scaled up on high-DPI screens while the CSS size stays unchanged.

There is no resize observer in v1: when the container changes size, `destroy()` the instance and create a new one.

## Navigation (0-based)

```ts
preview.goToSlide(0); // scroll to the first slide
preview.goToSlide(count - 1);
preview.nextSlide(); // from the most visible slide
preview.prevSlide();
preview.slideCount; // 0 before load completes / after destroy
```

Mixed boundaries — intentional, read carefully:

- A non-integer index **throws `TypeError`** (programming error, fail-fast).
- Before the load completes or after `destroy()`, navigation is a **no-op** (placeholders do not exist yet).
- After the load completes, an out-of-range index **throws `RangeError`**.

Hidden slides are all rendered and counted; indices follow the original deck order (nothing is skipped).

## Encrypted files

```ts
const preview = createPptPreview(container!, {
  source: encryptedFile,
  password: "open-password", // PowerPoint "password to open" (MS-OFFCRYPTO Agile)
  ...
});
```

Decryption uses WebCrypto on the main thread — it requires a secure context (HTTPS or localhost) and typically adds ~1s before parsing. A missing or wrong password reports `PASSWORD_PROTECTED` (the message distinguishes the two cases). ZIP-entry encryption is not covered.

## Error codes

| Code                 | Meaning                                                                                                         |
| -------------------- | --------------------------------------------------------------------------------------------------------------- |
| `UNSUPPORTED`        | Not a `.pptx` (unrecognized container), or a limitation a password cannot fix (incl. resource budgets exceeded) |
| `LEGACY_FORMAT`      | Legacy binary `.ppt` — re-save as `.pptx`                                                                       |
| `PASSWORD_PROTECTED` | Encrypted file: missing or wrong password (message distinguishes)                                               |
| `CORRUPT`            | Structurally corrupted content (only where stably distinguishable)                                              |
| `UNKNOWN`            | Source read, Worker/WASM startup, CSP or asset-URL failures — the file is not at fault                          |

## Destroy

```ts
preview.destroy();
```

Idempotent. Disconnects observers, expires in-flight renders, zeroes canvases, removes the injected DOM and terminates the engine's parser worker. All callbacks go silent after destroy. If called while the engine is still loading, the stale result is voided and the engine instance is destroyed as soon as it arrives (the upstream `load()` has no abort signal — this is not an immediate cancel).

## SSR

The module can be imported on the server, but `createPptPreview` must only run in a browser environment with `Worker` and `WebCrypto` available (guard it in `onMounted` / `useEffect`).
