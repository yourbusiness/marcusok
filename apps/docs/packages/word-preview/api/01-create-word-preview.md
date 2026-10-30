# createWordPreview

The single high-level entry. Renders a `.docx` into a container with format sniffing, error normalization, hardened defaults and post-render hardening.

> **Naming note:** `@marcusok/excel-preview` exports `createPreview`; this package exports `createWordPreview` deliberately — importing both on one page stays unambiguous.

```ts
import { createWordPreview } from "@marcusok/word-preview";

function createWordPreview(
  container: HTMLElement,
  options: WordPreviewOptions,
): WordPreviewInstance;
```

## WordPreviewOptions

| Option           | Type                                        | Default | Description                                                                                                                                                                                                                                                                                                                                                                                |
| ---------------- | ------------------------------------------- | ------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `source`         | `File \| Blob \| Uint8Array \| ArrayBuffer` | —       | The document bytes.                                                                                                                                                                                                                                                                                                                                                                        |
| `renderOptions`  | `WordRenderOptions`                         | —       | Whitelisted engine options (below).                                                                                                                                                                                                                                                                                                                                                        |
| `styleNonce`     | `string`                                    | —       | CSP nonce for injected `<style>` nodes. Implemented by node replacement (CSP rejects at insertion time — patching the attribute later would not work). See [Limits → CSP](../guide/02-limits#csp-requirements).                                                                                                                                                                            |
| `allowAltChunks` | `boolean`                                   | `false` | Render altChunks (embedded HTML). Off by default — the engine uses unsandboxed `srcdoc` iframes. Trusted documents only; a `sandbox=""` attribute is added after rendering (residual race, see [Limits](../guide/02-limits)).                                                                                                                                                              |
| `onRendered`     | `(info) => void`                            | —       | Called once per successful render. `info.pageCount` counts `section` elements (meaningful only under the default pagination options); `info.durationMs` is measured from the `createWordPreview` call — it covers source reading, format sniffing and the **full** render including image/font loading (the engine awaits all resource tasks before resolving), not just DOM construction. |
| `onError`        | `(error) => void`                           | —       | Fatal errors. Destroy-triggered cancellation is silent and never reaches this callback.                                                                                                                                                                                                                                                                                                    |

Errors carry a `code`:

| Code                 | Meaning                                                                                                                                           |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| `UNSUPPORTED`        | Not a ZIP container, or a ZIP missing required Word parts (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml`) — i.e. not a Word document. |
| `LEGACY_FORMAT`      | OLE2 container: legacy `.doc` **or** encrypted `.docx` (same magic bytes; both unsupported).                                                      |
| `PASSWORD_PROTECTED` | ZIP general-purpose encryption bit set (the rare directly-encrypted form).                                                                        |
| `CORRUPT`            | Broken ZIP structure, or required parts present but parsing/rendering failed (`cause` keeps the original error).                                  |
| `UNKNOWN`            | Anything else (e.g. the source `Blob` could not be read).                                                                                         |

## WordRenderOptions

The engine's 20 options split into **13 passable** + 7 taken over by this package. Passing the taken-over keys does nothing — they are overwritten.

**Passable:**

`breakPages`, `inWrapper`, `hideWrapperOnPrint`, `ignoreWidth`, `ignoreHeight`, `ignoreFonts`, `ignoreLastRenderedPageBreak` (package default `false`, engine default `true` — set `true` to restore engine behavior), `renderHeaders`, `renderFooters`, `renderFootnotes`, `renderEndnotes`, `renderComments`, `renderChanges`

All are optional booleans; unset keys use the engine defaults. `pageCount` in `onRendered` is meaningful only under the default pagination options (`breakPages: true`, `inWrapper: true`).

**Taken over (not in the type — 7 keys):**

| Key                                                   | Value                 | Why                                                                                            |
| ----------------------------------------------------- | --------------------- | ---------------------------------------------------------------------------------------------- |
| `className`                                           | `docx-<instanceSeq>`  | Unique per instance — isolates the engine's global document CSS between instances.             |
| `useBase64URL`                                        | `true`                | The engine never revokes blob URLs; base64 avoids the leak (memory cost documented in Limits). |
| `renderAltChunks`                                     | from `allowAltChunks` | Security default; see above.                                                                   |
| `h` / `experimental` / `debug` / `trimXmlDeclaration` | —                     | Experimental hooks and internals — not part of the public surface.                             |

## WordPreviewInstance

```ts
interface WordPreviewInstance {
  destroy(): void;
}
```

`destroy()` synchronously detaches the internal wrapper (the user container stays a plain mount point), drops any in-flight render result, and silences all callbacks. The engine has no cancellation API (0 abort paths upstream) — the synchronous DOM-build segment of an in-flight render cannot be interrupted, but its output is discarded. Idempotent; the container is immediately reusable by a new instance.
