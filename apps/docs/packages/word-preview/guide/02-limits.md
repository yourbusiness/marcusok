# Limits & Security Notes

This package is a thin, hardened wrapper around the [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) engine. The engine's constraints are inherited as-is; this page lists them honestly, together with the wrapper's own hardening and its residual risks.

## Rendering

- **Main-thread rendering.** The engine builds DOM on the main thread and has no API for yielding it; a large document blocks the UI for the duration of the DOM build. There is no worker mode (the engine's internal model is not structured-clone-safe — an upstream constraint, not a choice). Baseline measured in a real Chromium (six runs, 2026-10-22): 1000 paragraphs + a 100×5 table render in 16–28 ms (median ~20 ms), including the full `createWordPreview` cycle; your mileage scales with document size.
- **No real-time pagination.** Pages break only at explicit page breaks (`<w:br type="page"/>`), section breaks, and `lastRenderedPageBreak` markers saved by the editor. The engine does not re-flow text, so rendered page counts can differ from what Word would compute. This package defaults `ignoreLastRenderedPageBreak: false` to follow the saved markers (the engine's own default ignores them).
- **TOC fields do not render.**
- **Broken images/fonts fail silently.** The engine awaits its image/font load tasks with `Promise.allSettled` — a missing or corrupt image (or an embedded font that fails to decode) is dropped without an error: rendering still "succeeds", `onRendered` fires, and the bad resource is simply absent from the output. `onRendered` firing does not mean every resource landed.
- **No input-size cap.** A docx is untrusted input; there is no configurable limit on source bytes or decompressed size in this first version (no zip-bomb protection). Only preview files you would otherwise trust to open.
- **Containers must live in the main document.** Rendering into an `iframe` document or a ShadowRoot is not verified and not supported in this version (the engine reads global `document`, `Range`, `getComputedStyle`).

## Formats

- **`.docx` / `.docm` only.** Legacy `.doc` (OLE2) is unsupported.
- **Encrypted files.** The standard "encrypted docx" (Agile AES) is an OLE2 container — the same magic bytes as `.doc`; both report `LEGACY_FORMAT` with a message saying exactly that. The rare directly-encrypted ZIP form is detected by the ZIP general-purpose bit and reports `PASSWORD_PROTECTED`.

## Security hardening (what the wrapper does)

- **altChunks are off by default.** The engine renders altChunks (embedded HTML parts) into `<iframe srcdoc>` **without a `sandbox` attribute**, and its own default is `renderAltChunks: true` — untrusted documents could execute script on render. This package flips the default to `false`. Opting in via `allowAltChunks: true` is for trusted documents only; the wrapper then adds `sandbox=""` after rendering — note this is a post-render attribute assignment and the race with the iframe's first load is not guaranteed to be won in every browser. Content inside the iframe is also **outside** the hyperlink allowlist below.
- **Hyperlink scheme allowlist.** The engine assigns external relationship targets to `href` verbatim, with no scheme validation — a crafted docx can produce `<a href="javascript:…">`. After rendering, the wrapper parses every `a[href]` with the URL API (handling case, leading whitespace/control chars, `//host` forms) and keeps only `http:`, `https:`, `mailto:`, `tel:` plus anchor/relative forms; everything else gets its `href` removed (link text is kept). There is no opt-out — there is no legitimate case for `javascript:` links in a document preview.
- **Format sniffing before the engine.** Magic bytes, the ZIP encryption bit, and required-part checks run first, so hostile or malformed inputs get readable errors (`UNSUPPORTED` / `CORRUPT`) instead of the engine's null-deref stack traces.

## Residual risks (what the wrapper does not do)

- **CSS injection.** Document-controlled strings (numbering text, font names) end up inside the injected `<style>` CSS text. The wrapper's per-instance class name isolates _legitimate_ styles between instances, but a malicious payload can still close a rule and inject global selectors (e.g. `body { … }`) — verified in a real Chromium. Modern browsers do not execute script from CSS; the impact is UI spoofing and external `@import`/`url()` requests. Mitigate with CSP (`style-src`, `img-src`, `font-src`). No CSS sanitization is attempted — it would break the engine's legitimate base64 images/fonts.
- **Same-name embedded fonts across instances.** `@font-face` declarations are global by name; two documents embedding the same font name with different content — the later render wins.

## CSP requirements

`styleNonce` only covers the injected `<style>` **elements**. The engine also uses inline `style` **attributes** extensively (page size/margins, image positioning), and with `useBase64URL: true` images and fonts are `data:` URLs. A strict CSP therefore needs all of:

- `style-src 'nonce-…'` (or `'unsafe-inline'`) — for the injected styles
- `style-src-attr 'unsafe-inline'` (or an overall `style-src 'unsafe-inline'`) — for the inline attributes
- `img-src data:` and `font-src data:` — for the base64 resources

## Memory

`useBase64URL: true` (locked) means images and embedded fonts live in the DOM as base64 — roughly +33% over the raw bytes, resident as long as the preview is mounted. For image-heavy documents this is heavier than a blob-URL approach; the trade avoids the engine's unrevoked blob-URL leak.

## SSR

Importing the package is safe in Node (the module top level touches no DOM globals — verified). Only the `createWordPreview` call needs a DOM; call it client-side.

## Multi-instance

Each instance renders under a unique `docx-<n>` class name, isolating the engine's per-document CSS (default styles, document styles, numbering, theme variables). Residue: the `@font-face` name collision above. Verified in a real browser — two documents with different numbering keep their own numbering. One edge caveat: the instance counter lives inside each loaded copy of this library — loading two copies of `@marcusok/word-preview` on the same page (e.g. two different versions side by side) restarts the counter in each copy and the class names can collide; that layout is not supported.
