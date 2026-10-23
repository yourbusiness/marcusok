# @marcusok/word-preview

Read-only `.docx` preview in the browser with visual fidelity — pagination, headers/footers, footnotes/endnotes, comments and tracked changes render the way Word saved them. Rendering is powered by [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) (Apache-2.0) bundled into this package's `dist` together with its ZIP dependency (jszip) — **zero transitive dependencies** for consumers.

```ts
import { createWordPreview } from "@marcusok/word-preview";

const preview = createWordPreview(container, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onRendered: ({ pageCount, durationMs }) => console.log(pageCount, durationMs),
  onError: (e) => console.error(e.code, e.message),
});

preview.destroy();
```

## Why the wrapper (what you get over raw docx-preview)

| Concern                   | Raw docx-preview                                                                                | This package                                                                                                                                                     |
| ------------------------- | ----------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Bad input                 | A ZIP that isn't a docx dies with `Cannot read properties of null (reading 'find')`             | Sniffed before the engine: magic bytes + ZIP encryption bit + required parts → readable `UNSUPPORTED` / `LEGACY_FORMAT` / `PASSWORD_PROTECTED` / `CORRUPT` codes |
| altChunk script execution | On by default (`renderAltChunks: true`), rendered as unsandboxed `<iframe srcdoc>`              | Off by default; opting in (`allowAltChunks`) adds `sandbox=""` after rendering                                                                                   |
| Hyperlink XSS             | External relationship targets are assigned to `href` verbatim (`javascript:` survives)          | Post-render scheme allowlist (`http`/`https`/`mailto`/`tel` + anchors/relative), URL-API parsed — everything else loses its `href`                               |
| Multi-instance styles     | Fixed `docx` class name — the later document's numbering/theme CSS overwrites the earlier one   | Unique per-instance class name; instances never clobber each other                                                                                               |
| Blob URL leak             | `useBase64URL: false` default and zero `revokeObjectURL` calls upstream                         | `useBase64URL: true` (base64 lives in the DOM; ~+33% memory for images/fonts, no leak)                                                                           |
| destroy() while rendering | The engine clears and appends into your container _after_ its awaits — a late render refills it | Renders into an internal wrapper; `destroy()` detaches it synchronously, late results are dropped, callbacks go silent                                           |
| CSP                       | Injected `<style>` blocked under strict CSP                                                     | `styleNonce` rebuilds the style nodes with your nonce (node replacement — patching the attribute later would not work)                                           |

## Limits (inherited from the engine, stated honestly)

- **Main-thread rendering** — no worker mode; large documents block the UI for the DOM build.
- **No real-time pagination** — pages break at explicit breaks and `lastRenderedPageBreak` markers only; TOC fields do not render.
- **`.docx` / `.docm` only** — legacy `.doc` and encrypted docx are both OLE2 and both rejected; the directly-encrypted ZIP form is rejected too.
- **CSS injection surface** — document-controlled strings reach the injected `<style>` text; the class-name isolation covers legitimate styles only. Use CSP for hostile inputs (`style-src`, `img-src data:`, `font-src data:`).
- **Containers must be in the main document** — iframe documents / ShadowRoots are not verified.

Full details in the [docs site limits page](https://yourbusiness.github.io/marcusok/packages/word-preview/guide/02-limits).

## Performance

- **Bundle**: ~289 KB / ~62 KB gzipped (unminified, matching this repo's convention), self-contained.
- **happy-dom regression baseline** (DOM construction only, no real layout): 1000 paragraphs + a 100×5 table renders in ~96 ms median on the dev machine (single cold call incl. JIT); the `performance.test.ts` SLA is 150 ms.
- **Real Chromium baseline** (measured 2026-10-22, six runs, full render cycle through `createWordPreview`, `onRendered.durationMs` semantics): 1000 paragraphs + a 100×5 table ≈ 16–28 ms (median ~20 ms; 1500 `<p>` + 100 table rows produced). Word-sized documents (tens of pages, no image floods) stay comfortably interactive.

## Install

```bash
pnpm add @marcusok/word-preview
```

## API

One entry point — `createWordPreview(container, options) → { destroy() }`. Options:

- `source` — `File | Blob | Uint8Array | ArrayBuffer`
- `renderOptions` — 13 whitelisted engine options (`breakPages`, `inWrapper`, `hideWrapperOnPrint`, `ignoreWidth`, `ignoreHeight`, `ignoreFonts`, `ignoreLastRenderedPageBreak`, `renderHeaders`, `renderFooters`, `renderFootnotes`, `renderEndnotes`, `renderComments`, `renderChanges`). `className`, `useBase64URL`, `renderAltChunks` and the experimental hooks are managed by this package.
- `styleNonce` — CSP nonce for injected `<style>` (rebuilt via node replacement).
- `allowAltChunks` — opt in to altChunk rendering (trusted documents only).
- `onRendered(info)` — `{ pageCount, durationMs }`; `durationMs` is measured from the `createWordPreview` call (source read + sniffing + render incl. image/font loading), `pageCount` counts `section` elements (meaningful under default pagination options).
- `onError(error)` — `{ code, message, cause? }`.

SSR note: importing the package is safe in Node (no DOM globals touched at module top level); only the render call needs a DOM.

## License

MIT. Third-party notices (docx-preview under Apache-2.0, jszip/pako under MIT) are in [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md); both projects' license headers are preserved in the built `dist/index.js`.
