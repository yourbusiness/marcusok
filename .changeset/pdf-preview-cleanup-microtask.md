---
"@marcusok/pdf-preview": patch
---

Defer the static `TextLayer.cleanup()` call to a microtask when the last preview instance is destroyed. `teardown()` cancels in-flight text layers and calls cleanup in the same synchronous block, but pdf.js only removes a cancelled instance from its module-level `#pendingTextLayers` set in a settle microtask — so the synchronous `cleanup()` saw a non-empty set, returned early, and nothing ever called it again: the hidden measuring canvases and the ascent cache stayed alive until the page unloaded. The deferred callback re-checks the active-instance count so an instance created in the same frame is never cleaned.
