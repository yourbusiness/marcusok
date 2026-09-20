---
"@marcusok/excel-exporter": patch
---

Fix the WASM loader getting stuck in its error state after a late success: `initWasm` keeps a single module-level in-flight promise, so on a slow network the underlying load could still succeed after all retried attempts had timed out — the engine was then initialized, yet every later export kept failing with "WASM load previously failed" until `configureWasm()` was called. A late success now flips the loader back to ready.
