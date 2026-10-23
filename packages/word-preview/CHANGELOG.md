# @marcusok/word-preview

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
