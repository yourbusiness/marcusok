---
"@marcusok/pdf-preview": patch
---

Cap high-DPR canvas bitmaps at the pdf.js viewer's pixel budget (5,242,880) to avoid multi-hundred-MB page bitmaps; re-probe the scroll root after all placeholders are inserted and ignore percentage `max-height` wrappers; re-queue pages whose stale release event arrives after they re-entered the viewport; fix `fit-page` padding for content-box containers and warn when the height is unmeasurable; clean pdf.js text-layer caches when the last instance is destroyed; sniff before copying typed sources; and center oversized pages with `safe center` so their left edge stays reachable.
