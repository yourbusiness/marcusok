---
"@marcusok/excel-exporter": patch
---

- fix: the Workbook path now normalizes non-primitive cell values exactly like the stream and SheetJS paths — plain objects are written as JSON strings, `Date`s as ISO strings, bigints as decimal strings — so a dataset crossing the 50k-row threshold (or exported before/after a degradation) keeps identical cell content. Previously objects landed as `"[object Object]"` and dates as locale-dependent long text on the Workbook path only.
  - fix: `toStr` now handles `symbol`/`function` values explicitly (`JSON.stringify` returns `undefined` for them), so such cells receive a visible string instead of `undefined`.
  - fix: the worker no longer re-runs `initWasm` and re-reports the `init` phase on every export: a URL-typed `wasmUrl` arrives via structured clone as a fresh object each message, so the old reference comparison never matched. Initialization is now tracked by the URL's string form plus an explicit ready flag. Note `initWasm` is idempotent — a `wasmUrl` change after the first successful init does not take effect (the stale comment claiming otherwise is corrected).
  - fix: when the browser Worker route fails (missing/404 `workerUrl`, WASM init error in the worker, timeout), `exportExcel` now retries on the main thread with modern-xlsx first (styles preserved; the ≥50k-row tier uses the WASM-free fast stream) and only degrades to the style-less SheetJS fallback if that retry also fails. The `onProgress` 0→1 contract (each exactly once) is unchanged.
  - tests: 91 → 94 (CI runs 90): cross-path value-normalization regression (Workbook vs stream read back cell-by-cell), worker-failure → main-thread retry (workbook and stream tiers), and the full worker → main-thread → SheetJS double-failure chain.
