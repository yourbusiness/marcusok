---
"@marcusok/word-preview": patch
---

Create altChunk iframes with `sandbox=""` before `srcdoc` is assigned (post-render patching ran after Chromium had already executed the script), resolve the main document part from `_rels/.rels` instead of hard-coding `word/document.xml`, classify encrypted ZIPs whose central directory carries the encryption bit as `PASSWORD_PROTECTED`, require a full ZIP signature for `PK`-prefixed text, rewrite the opaque missing-relationship `TypeError` into an actionable message, declare the JSZip import side effect (`sideEffects: true`), and complete the third-party license notices.
