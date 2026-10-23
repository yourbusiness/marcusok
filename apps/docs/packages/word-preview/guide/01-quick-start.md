# Quick Start

Requirement: a browser DOM. Importing the package is SSR-safe (the module top level touches no DOM globals — only the render call needs a `document`), but `createWordPreview` itself must run in the browser.

## 1. Install

```bash
pnpm add @marcusok/word-preview
```

Nothing else to wire: the rendering engine (docx-preview) and its ZIP dependency (jszip) are bundled into the package's `dist` at build time. No WASM, no worker scripts, no bundler plugins.

## 2. Mount a preview

```ts
import { createWordPreview } from "@marcusok/word-preview";

const preview = createWordPreview(container, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onRendered: ({ pageCount, durationMs }) => {
    console.log(`${pageCount} sections rendered in ${durationMs}ms`);
  },
  onError: (e) => {
    // e.code: UNSUPPORTED | LEGACY_FORMAT | PASSWORD_PROTECTED | CORRUPT | UNKNOWN
    console.error(e.code, e.message);
  },
});
```

The container is only a mount point — the preview renders into an internal wrapper, so `destroy()` can detach it synchronously even while a render is still in flight.

## 3. Unmount

```ts
preview.destroy();
```

`destroy()` takes effect immediately: the rendered DOM is detached, a render that resolves afterwards is dropped, and all callbacks go silent. Calling it twice is safe. The same container can be reused by a new instance right away.

## Where the defaults differ from the engine

Three defaults intentionally differ from docx-preview's own defaults, tuned for preview use:

| Option                        | Engine default | This package                   | Why                                                                                                                                              |
| ----------------------------- | -------------- | ------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `useBase64URL`                | `false`        | `true` (locked)                | The engine never revokes blob URLs (0 `revokeObjectURL` calls upstream); base64 avoids the leak. Cost: ~+33% memory for images/fonts in the DOM. |
| `ignoreLastRenderedPageBreak` | `true`         | `false` (overridable)          | Follows Word's saved pagination more closely. The marker is written by the editor, not computed by layout — it can differ from true page breaks. |
| `renderAltChunks`             | `true`         | `false` (see `allowAltChunks`) | The engine renders altChunks as unsandboxed `srcdoc` iframes — a script-execution vector for untrusted documents.                                |

Everything else passable lives in `renderOptions` (see the [API](../api/01-create-word-preview#wordrenderoptions)).

## Next steps

- [Limits](./02-limits) — main-thread rendering, CSP requirements, CSS injection, multi-instance residue.
- Try it with your own files in the [play](/play).
