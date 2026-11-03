---
"@marcusok/pdf-preview": patch
---

Call the pdf.js `TextLayer.cleanup()` static with `TextLayer` as the receiver. The method reads `this.#pendingTextLayers`; the previous detached invocation threw `TypeError` on every last-instance destroy and was swallowed by the catch, so the global text-layer measurement cache was never actually cleaned.
