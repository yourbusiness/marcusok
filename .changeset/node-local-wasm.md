---
"@marcusok/excel-exporter": patch
"@marcusok/excel-preview": patch
---

Support local WASM files in Node: `configureWasm({ wasmUrl })` now accepts a plain filesystem path or a `file:` URL in addition to HTTP(S) URLs. The binary is read synchronously (matching zero-config auto-init), and a missing/unreadable configured file fails with the real fs/compile reason instead of a misleading `fetch failed`.
