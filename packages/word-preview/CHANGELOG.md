# @marcusok/word-preview

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
