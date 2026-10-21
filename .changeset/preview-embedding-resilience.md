---
"@marcusok/excel-preview": minor
---

Improve embedding and worker robustness: styles are injected into the container's ShadowRoot (with an optional CSP `styleNonce`) and the viewport window is defensively capped when CSS did not apply; the parse worker now uses a FIFO queue where timeouts measure processing time and do not reject queued parses; `terminateParseWorker()` clears the `workerBroken` fallback; non-integer/NaN `setSheet()` indexes report through `onError` instead of throwing; and `onParsed`/`onError` callback exceptions are isolated.
