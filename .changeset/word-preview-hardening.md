---
"@marcusok/word-preview": patch
---

Clear the detached render DOM on destroy and all late-render paths, enforce the documented runtime takeover of `h`/`experimental`/`debug`/`trimXmlDeclaration`, accept cross-realm `Uint8Array` sources and cross-realm thenable callback rejections, and correct the `breakPages`/`pageCount` JSDoc.
