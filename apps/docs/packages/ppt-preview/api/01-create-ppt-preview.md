# API: `createPptPreview`

```ts
import { createPptPreview } from "@marcusok/ppt-preview";

function createPptPreview(
  container: HTMLElement,
  options: PptPreviewOptions,
): PptPreviewInstance;
```

Mounts a read-only `.pptx` preview into `container` and returns a handle. All public types are self-contained in the package (nothing is re-exported from the engine) — see `dist/index.d.ts`.

## `PptPreviewOptions`

| Option                | Type                                        | Default                   | Description                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| --------------------- | ------------------------------------------- | ------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `source`              | `File \| Blob \| Uint8Array \| ArrayBuffer` | —                         | **Required.** Bytes are defensively copied before the engine runs (the parser transfers its input buffer); the caller's original bytes stay usable.                                                                                                                                                                                                                                                                                                                                               |
| `password`            | `string`                                    | —                         | Open-password for MS-OFFCRYPTO Agile encrypted files. Decryption runs on the main thread (WebCrypto, secure context, ~+1s). ZIP-entry encryption is not covered.                                                                                                                                                                                                                                                                                                                                  |
| `slideWidth`          | `number`                                    | `960`                     | CSS pixel width of each slide (fractional values are floored); height follows the deck's aspect ratio. Must be a positive finite number — otherwise a synchronous `TypeError`.                                                                                                                                                                                                                                                                                                                    |
| `dpr`                 | `number`                                    | `window.devicePixelRatio` | Canvas bitmap scale. Must be a positive finite number — otherwise a synchronous `TypeError`.                                                                                                                                                                                                                                                                                                                                                                                                      |
| `enableTextSelection` | `boolean`                                   | `true`                    | Text layer (select/copy). Runs are collected in the same render pass as the canvas paint — no double rendering. Positioning uses CSSOM; see the CSP notes in [Limits](/packages/ppt-preview/guide/02-limits).                                                                                                                                                                                                                                                                                     |
| `useGoogleFonts`      | `boolean`                                   | `false`                   | Allow Google Fonts substitution fetches for missing fonts. Off by default: no implicit external requests.                                                                                                                                                                                                                                                                                                                                                                                         |
| `wasmUrl`             | `string \| URL`                             | engine anchor             | Explicit WASM asset URL escape hatch for self-hosting / rewritten asset paths. Default requires zero config under bundlers (the asset next to the entry is emitted automatically).                                                                                                                                                                                                                                                                                                                |
| `styleNonce`          | `string`                                    | —                         | CSP nonce applied to **this package's** injected `<style>` (style-src). CSSOM-written positioning is not governed by style-src-attr in Chromium (see Limits).                                                                                                                                                                                                                                                                                                                                     |
| `renderOptions`       | `PptRenderOptions`                          | —                         | Allowlist passthrough (below).                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `onRendered`          | `(info) => void`                            | —                         | Fired once, by the **first successfully rendered slide**. `slideCount` is the raw deck total (hidden slides included); `firstSlideMs` is that slide's canvas + text-layer time; `durationMs` covers everything since `createPptPreview` (read, sniff, decrypt, **full-deck preflight**, render). Not fired when every slide fails, the deck is empty, or the container is hidden/zero-sized at creation (placeholders never enter the render band; it starts once the container becomes visible). |
| `onError`             | `(error) => void`                           | —                         | Load failures and per-slide render failures, normalized to the public error codes. Destroy-triggered cancellation is silent. User callback exceptions are isolated (`console.warn`), never reported as preview errors.                                                                                                                                                                                                                                                                            |

### `PptRenderOptions`

| Option            | Type                                             | Default  | Description                                                                                                                                                                                 |
| ----------------- | ------------------------------------------------ | -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cjkFallback`     | `"auto" \| "kr" \| "sc" \| "tc" \| "hk" \| "jp"` | `"auto"` | Substitute-glyph target for missing CJK fonts.                                                                                                                                              |
| `resourceLimits`  | `PptResourceLimits`                              | engine   | Archive budgets: `maxArchiveEntryBytes`, `maxTotalInflatedBytes`, `maxArchiveEntries` (`number \| null`; `null` disables a limit). Exceeding a budget reports `UNSUPPORTED`, not `CORRUPT`. |
| `workerTimeoutMs` | `number`                                         | engine   | Parser worker per-response timeout; bounds both normal loading and the worst-case wait during destroy-while-loading.                                                                        |

Parameters the engine supports but this package intentionally does **not** expose: `mode` (fixed `"main"`), `progressiveLayout` (off — stable slide count and placeholders), `math`/`threeD`/`regionMap`/`chartEx`/`tiff` (optional renderers, not injected in v1), and the deprecated `maxZipEntryBytes` (use `resourceLimits.maxArchiveEntryBytes`).

## `PptPreviewInstance`

### `destroy(): void`

Idempotent. Disconnects observers, expires in-flight renders (the engine has no per-slide cancel handle — completion paths clean up), zeroes canvases, removes injected DOM, destroys the engine (terminating the parser worker). Callbacks go silent. When called while loading, the stale load is voided and the engine instance is destroyed on arrival (not an immediate cancel — upstream `load()` has no abort signal).

### `goToSlide(index: number): void`

Scrolls to a slide (0-based, original deck order, hidden slides included). When a scroll container is identified, only that container is scrolled. When the container itself does not scroll (plain document flow), it falls back to `window`, so the outer page moves.

Mixed boundaries (intentional):

| Situation                                 | Behavior                                                       |
| ----------------------------------------- | -------------------------------------------------------------- |
| Non-integer index                         | throws `TypeError` (checked before state — also after destroy) |
| Before load completes / after `destroy()` | no-op                                                          |
| Loaded, out of range                      | throws `RangeError`                                            |

### `nextSlide(): void` / `prevSlide(): void`

Step from the **most visible slide** (measured on demand — manual scrolling is always respected), not from the last `goToSlide`. No-ops at the boundaries, before load, and after destroy.

### `get slideCount(): number`

Raw deck total (hidden slides included). `0` before the load completes, after `destroy()`, and for empty decks.

## Error codes

```ts
type PptPreviewErrorCode =
  | "UNSUPPORTED" // not .pptx / engine-unsupported features / resource budgets
  | "LEGACY_FORMAT" // legacy binary .ppt
  | "PASSWORD_PROTECTED" // encrypted: missing or wrong password (message distinguishes)
  | "CORRUPT" // stably-distinguishable structural corruption
  | "UNKNOWN"; // source read / Worker / WASM / CSP / asset URL / unknown traps
```

Engine and source-read failures carry `cause` (the original error) for diagnostics; errors constructed locally by format sniffing or an invalid deck size may omit it.
