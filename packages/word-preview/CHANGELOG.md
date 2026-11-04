# @marcusok/word-preview

## 0.1.7

### Patch Changes

- c5b85b8: Three fixes: (1) the normalized "render failed" message for engine null-reference crashes now phrases the missing-relationship-part cause as _a common cause_ instead of asserting it — other null crashes inside the engine produced the same message and were misattributed to a broken file. (2) Anchors with leading C0/space characters (`href=" #section"`) are preserved on non-http(s) pages (file://, custom protocols), matching how browsers strip those characters before resolving the href. (3) The package re-enables the repo-wide `noUncheckedIndexedAccess` TypeScript check (the ZIP flag sniff in production code now reads the flag byte defensively).

## 0.1.6

### Patch Changes

- ca0d55a: Keep `#anchor` hyperlinks on non-http(s) pages (e.g. `file://` reports, Electron): fragment-only hrefs are now always allowed instead of inheriting the page protocol and being stripped by the scheme allowlist. Also refresh the published JSDoc for `allowAltChunks` (the sandbox attribute is set at iframe creation, before `srcdoc` is assigned — not patched after render) and add a full-pipeline performance gate covering `createWordPreview` (sniffing + second zip directory parse + render), which previously had no regression protection.

## 0.1.5

### Patch Changes

- 90fe3d4: Sniff the main part's OPC content type so xlsx/pptx inputs report `UNSUPPORTED` instead of a misleading `CORRUPT`; clean up the engine's `src="null"` / `url(null)` placeholders so missing images/fonts no longer trigger `/null` requests or broken icons; and normalize null-dereference render failures across V8, Firefox and Safari messages into the actionable "missing relationship part" error.

## 0.1.4

### Patch Changes

- f2e7919: Create altChunk iframes with `sandbox=""` before `srcdoc` is assigned (post-render patching ran after Chromium had already executed the script), resolve the main document part from `_rels/.rels` instead of hard-coding `word/document.xml`, classify encrypted ZIPs whose central directory carries the encryption bit as `PASSWORD_PROTECTED`, require a full ZIP signature for `PK`-prefixed text, rewrite the opaque missing-relationship `TypeError` into an actionable message, declare the JSZip import side effect (`sideEffects: true`), and complete the third-party license notices.

## 0.1.3

### Patch Changes

- 6a0be9e: Clear the detached render DOM on destroy and all late-render paths, enforce the documented runtime takeover of `h`/`experimental`/`debug`/`trimXmlDeclaration`, accept cross-realm `Uint8Array` sources and cross-realm thenable callback rejections, and correct the `breakPages`/`pageCount` JSDoc.

## 0.1.2

### Patch Changes

- f9d1ef5: Clean up comment highlights registered after an in-flight render is destroyed.

## 0.1.1

### Patch Changes

- 71229de: Clean up comment highlight registrations on destroy and failed renders, and isolate rejected promises from user callbacks.

## 0.1.0

### Minor Changes

- e7165b3: Initial release: read-only docx preview in the browser with visual fidelity.

  - Single entry `createWordPreview(container, options)` with format sniffing before the engine (magic bytes + ZIP encryption bit + required-part checks) and normalized error codes (`UNSUPPORTED` / `LEGACY_FORMAT` / `PASSWORD_PROTECTED` / `CORRUPT` / `UNKNOWN`)
  - Hardened defaults over docx-preview: altChunks off (unsandboxed `srcdoc` iframes), post-render hyperlink scheme allowlist (http/https/mailto/tel + anchors/relative), `useBase64URL: true` (avoids the engine's unrevoked blob URLs), `ignoreLastRenderedPageBreak: false` (follows Word's saved pagination)
  - Per-instance unique style class names — multiple documents on one page never clobber each other's numbering/theme CSS
  - `destroy()` detaches the internal wrapper synchronously; late in-flight renders are dropped and callbacks go silent
  - Error paths (failed sniff or render) remove the internal wrapper — a failed preview leaves the user's container untouched
  - `styleNonce` support via style-node replacement (CSP-safe), CSP requirements documented
  - docx-preview 0.4.1 (Apache-2.0) + jszip 3.10.1 (MIT) bundled into dist — zero transitive dependencies for consumers
