# @marcusok/word-preview

Read-only `.docx` preview in the browser with visual fidelity — pagination, headers/footers, footnotes/endnotes, comments and tracked changes render the way Word saved them. Rendering is powered by [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) (bundled into this package's `dist` — zero transitive dependencies for consumers), wrapped with format sniffing, readable error codes, hardened security defaults and per-instance style isolation.

## Capabilities

| Capability        | Description                                                                                                                                                                                                                                    |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Formats           | `.docx` / `.docm` (ZIP/OOXML). Legacy `.doc` and encrypted `.docx` share the OLE2 magic bytes and both report `LEGACY_FORMAT`; the rare directly-encrypted ZIP form reports `PASSWORD_PROTECTED`                                               |
| Visual fidelity   | Page breaks (explicit `<w:br type="page"/>`, `lastRenderedPageBreak` markers, section breaks), headers/footers, footnotes/endnotes, comments, tracked changes, embedded images and fonts — as rendered by the docx-preview engine              |
| Format sniffing   | Magic bytes + ZIP encryption bit + required-part validation (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml`) run **before** the engine, so a non-Word ZIP gets a readable `UNSUPPORTED` error instead of a cryptic null dereference |
| Hardened defaults | `altChunks` (embedded HTML rendered in unsandboxed `srcdoc` iframes by the engine) are **off**; hyperlink `href`s outside the `http`/`https`/`mailto`/`tel` allowlist are stripped after rendering (link text kept)                            |
| Multi-instance    | Each instance renders under a unique class name, so two documents on one page never clobber each other's numbering/theme CSS (the engine's fixed class name would)                                                                             |
| CSP support       | `styleNonce` rebuilds the injected `<style>` nodes with your nonce (node replacement — CSP rejects at insertion time); see [Limits](./guide/02-limits) for the inline-style and `data:` requirements                                           |
| Cancellation      | `destroy()` detaches the internal wrapper synchronously — a render that resolves late writes into the detached node and is dropped; all callbacks go silent. Cancellation is not an error (no `onError`)                                       |
| Errors            | Code-carrying errors: `UNSUPPORTED`, `LEGACY_FORMAT`, `PASSWORD_PROTECTED`, `CORRUPT`, `UNKNOWN`                                                                                                                                               |

## Install

```bash
pnpm add @marcusok/word-preview
```

That single install brings everything: docx-preview and its ZIP dependency (jszip) are bundled into this package's `dist` at build time (both are devDependencies — consumers get **zero transitive dependencies**). See [Quick Start](/packages/word-preview/guide/01-quick-start).

## Quick example

```ts
import { createWordPreview } from "@marcusok/word-preview";

const preview = createWordPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onRendered: (info) => console.log(info.pageCount, info.durationMs),
  onError: (e) => console.error(e.code, e.message),
});

preview.destroy(); // unmount + neutralize any in-flight render
```

There is intentionally **no low-level data-model API**: the engine only guarantees `renderAsync` as stable (`WordDocument` is an opaque `any` stub upstream), so a `parse()`-style API cannot be honestly provided. This differs from [@marcusok/excel-preview](/packages/excel-preview/), which exposes `parseWorkbookBytes` — see its docs for that pattern.

<ClientOnly>
<WordPreviewDemo />
</ClientOnly>

## Not in scope

- **Real-time layout pagination** — pages break only at explicit page/section breaks and at `lastRenderedPageBreak` markers saved by the editor; the engine does not re-flow text.
- **TOC fields** — table-of-contents fields do not render.
- **Legacy `.doc`-style editing** — this is a viewer; nothing is edited or re-saved.
- **Worker parsing** — the engine produces DOM on the main thread by design; large documents block the UI for the duration of the DOM build (see [Limits](./guide/02-limits)).
