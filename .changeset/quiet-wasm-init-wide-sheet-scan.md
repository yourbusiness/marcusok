---
"@marcusok/excel-exporter": patch
"@marcusok/excel-preview": patch
---

- Quiet WASM init: the bundled loader now calls the engine with wasm-bindgen's official object form, so the `using deprecated parameters for the initialization function` console warning no longer appears on every WASM initialization (main thread and workers alike).
- `excel-preview`: wide-sheet horizontal scrolling is much cheaper — the renderer's per-frame window diff walks a sorted per-row cell index with binary search instead of scanning every cell of each visible row, so each scroll frame only touches cells inside the visible column window (measured 67x faster on a 60-row x 16384-column sheet; a performance baseline now guards this path). The README also documents the Vite dev-server `optimizeDeps.exclude` fix (without it the parse worker 404s and the session silently falls back to main-thread parsing), and the `destroy()` JSDoc now matches the implementation: aborting the active parse terminates the shared worker and re-dispatches queued parses to a fresh instance — it does not reject sibling requests.
