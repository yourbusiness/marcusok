---
"@marcusok/word-preview": patch
---

Keep `#anchor` hyperlinks on non-http(s) pages (e.g. `file://` reports, Electron): fragment-only hrefs are now always allowed instead of inheriting the page protocol and being stripped by the scheme allowlist. Also refresh the published JSDoc for `allowAltChunks` (the sandbox attribute is set at iframe creation, before `srcdoc` is assigned — not patched after render) and add a full-pipeline performance gate covering `createWordPreview` (sniffing + second zip directory parse + render), which previously had no regression protection.
