---
"@marcusok/excel-preview": minor
---

Add `terminateParseWorker()` — a public escape hatch that terminates the shared parse worker and rejects its in-flight parses (`noFallback`, no main-thread rerun; same policy as a worker timeout). Changing `parseWorkerUrl`/`wasmUrl` previously had no public way to recycle the already-created worker: the worker is a module-level singleton, and the shared loader's guidance suggested recreating the preview, which does not recreate the worker. After `terminateParseWorker()`, the next parse rebuilds a fresh worker from the current configuration. The shared loader's reconfiguration warnings (bundled from xlsx-core) now point at this escape hatch instead of the ineffective advice.
