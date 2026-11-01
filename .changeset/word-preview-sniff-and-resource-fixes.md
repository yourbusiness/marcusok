---
"@marcusok/word-preview": patch
---

Sniff the main part's OPC content type so xlsx/pptx inputs report `UNSUPPORTED` instead of a misleading `CORRUPT`; clean up the engine's `src="null"` / `url(null)` placeholders so missing images/fonts no longer trigger `/null` requests or broken icons; and normalize null-dereference render failures across V8, Firefox and Safari messages into the actionable "missing relationship part" error.
