---
"@marcusok/excel-exporter": patch
---

Isolate user callback errors: a throwing `onProgress`/`onPhase` no longer rejects the export or triggers a wasted degradation rebuild (previously, a throw from the trailing `onProgress(1)` after a successful Worker build re-built the whole workbook on the main thread). Callback exceptions are now logged via `console.warn` and the export completes normally, consistently across the main / worker / stream routes.
