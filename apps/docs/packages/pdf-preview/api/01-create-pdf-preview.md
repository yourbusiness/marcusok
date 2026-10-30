# createPdfPreview

```ts
import { createPdfPreview } from "@marcusok/pdf-preview";

const preview = createPdfPreview(container, options);
// preview: { destroy(): void }
```

Mounts a continuous-scroll, lazily rendered PDF preview into `container` (the container is only a mount point — the preview renders into an internal wrapper that `destroy()` detaches synchronously).

## PdfPreviewOptions

| Option                | Type                                        | Default       | Description                                                                                                                                                                                                                                                                                            |
| --------------------- | ------------------------------------------- | ------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `source`              | `File \| Blob \| Uint8Array \| ArrayBuffer` | —             | The PDF bytes. Read failures report `UNKNOWN`                                                                                                                                                                                                                                                          |
| `scale`               | `number \| "fit-width" \| "fit-page"`       | `"fit-width"` | Resolved once at creation. Numbers must be positive and strings must be `"fit-width"`/`"fit-page"` (a `TypeError` is thrown synchronously otherwise). Hidden containers (zero width) fall back to `1` + `console.warn`                                                                                 |
| `textLayer`           | `boolean`                                   | `true`        | Selectable/copyable transparent text layer over the canvas                                                                                                                                                                                                                                             |
| `workerSrc`           | `string`                                    | derived       | Worker script URL. Default: bundler-emitted anchor literal, with a `../dist/` runtime fallback for source-mode execution. Absolute or root-relative/relative URLs are accepted; relative ones resolve against the page base before cmaps/fonts derivation (see [Quick Start](../guide/01-quick-start)) |
| `cMapUrl`             | `string`                                    | derived       | CMap directory URL (trailing `/`). **Required for CJK documents** in bundler scenarios; missing cmaps degrade silently                                                                                                                                                                                 |
| `standardFontDataUrl` | `string`                                    | derived       | Standard-font directory URL (trailing `/`); needed when PDFs reference non-embedded standard fonts                                                                                                                                                                                                     |
| `wasmUrl`             | `string`                                    | —             | pdf.js wasm directory (not bundled; JPX/JBIG2/ICC features degrade without it — see [Limits](../guide/02-limits))                                                                                                                                                                                      |
| `iccUrl`              | `string`                                    | —             | ICC profile directory (not bundled; color management degrades without it)                                                                                                                                                                                                                              |
| `renderOptions`       | `PdfRenderOptions`                          | —             | Whitelist passthrough to `getDocument` (below)                                                                                                                                                                                                                                                         |
| `styleNonce`          | `string`                                    | —             | Nonce on the package's injected `<style>`. Does **not** cover pdf.js's own runtime font styles — see [Limits](../guide/02-limits#csp-requirements)                                                                                                                                                     |
| `onRendered`          | `(info: PdfRenderedInfo) => void`           | —             | Fires once, when the first **successfully** rendered page is painted (lazy pipeline — there is no "all pages done" event; if the very first page fails, the callback fires on the next page that succeeds)                                                                                             |
| `onError`             | `(error: PdfPreviewError) => void`          | —             | All failure paths. `destroy()`-triggered cancellations are silent (not errors)                                                                                                                                                                                                                         |

## PdfRenderOptions

The whitelist of `getDocument` parameters consumers may pass through:

| Option            | Type      | Description                                      |
| ----------------- | --------- | ------------------------------------------------ |
| `disableFontFace` | `boolean` | Skip loading embedded fonts (debug/fallback aid) |
| `useSystemFonts`  | `boolean` | Prefer system fonts over the standard-font files |

Everything else is taken over by the package: `data` (from `source`), `cMapUrl`/`cMapPacked`/`standardFontDataUrl`/`wasmUrl`/`iccUrl` (asset derivation + the options above), `enableXfa` (pinned `false`), verbosity (left at upstream defaults), and the url-only parameters (`disableAutoFetch`, `rangeChunkSize`) which are meaningless for in-memory sources. `isEvalSupported` no longer exists in pdf.js 6.x.

## PdfRenderedInfo

| Field         | Description                                                                                 |
| ------------- | ------------------------------------------------------------------------------------------- |
| `pageCount`   | Document total (`numPages`)                                                                 |
| `firstPageMs` | First painted page: canvas + text layer                                                     |
| `durationMs`  | From the `createPdfPreview` call to that first paint (source read + sniff + parse + render) |

## Error codes

| Code                 | When                                                                                                          |
| -------------------- | ------------------------------------------------------------------------------------------------------------- |
| `UNSUPPORTED`        | No `%PDF-` magic within the first 1024 bytes, or empty input                                                  |
| `PASSWORD_PROTECTED` | Encrypted PDF (upstream `PasswordException`; the message distinguishes needs-password from rejected-password) |
| `CORRUPT`            | Magic present but parsing/rendering failed (`InvalidPDFException` and unexpected errors; `cause` preserved)   |
| `UNKNOWN`            | Outside the pipeline — source read failure, worker/asset URL unreachable, or invalid asset URL configuration  |

## PdfPreviewInstance

```ts
interface PdfPreviewInstance {
  destroy(): void;
}
```

`destroy()` disconnects observers, cancels in-flight `RenderTask`s and text-layer streams, zeroes rendered canvases, calls `loadingTask.destroy()` (terminating the worker), and detaches the DOM. Idempotent; callbacks fall silent afterwards. The shared stylesheet singleton is intentionally kept (other instances may be live).
