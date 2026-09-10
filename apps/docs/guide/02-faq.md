# FAQ

### WASM 404 in the browser

With the default (zero-config) resolution this should not happen on Vite / webpack 5 — the bundler emits the shipped `modern-xlsx.wasm` as a hashed asset automatically. If you see a 404, you are likely overriding the URL (`configureWasm({ wasmUrl })` pointing at a wrong path) or using a bundler without `new URL(asset, import.meta.url)` support; point `configureWasm` at a URL your site actually serves, or copy the file out of this package's `dist/` into a static directory.

### Worker mode degrades to the main thread

The worker asset (`export.worker.js`) resolves automatically; degradation happens when the Worker route fails (e.g. a misconfigured `workerUrl` override 404s). The export then **retries on the main thread** (modern-xlsx keeps styles; the fast stream needs no WASM at all) — the style-less stream fallback is only the last resort when the main-thread retry also fails. Check for `[excel-exporter]` console warnings to find the reason.

### `result.error` is set although `success` is true

The export degraded to the style-less fast stream (WASM failed or is unsupported). Styles are stripped; headers and merges are preserved. The reason is in `result.error.message` and in the `[excel-exporter]` console warnings — usually a 404 wasm URL. See [fallback](/packages/excel-exporter/guide/08-fallback).

### Exporting 100k rows is very slow (>15s)

You are most likely on the `main` + `Workbook.toBuffer()` path, which has a cliff beyond ~55k rows. Keep `mode: "auto"` (~0.8s at 100k rows), or set `mode: "stream"` / `mode: "worker"` explicitly. See [auto mode](/packages/excel-exporter/guide/03-auto-mode).

### Styles do not apply in stream mode

Stream (v1) supports multi-row headers (`children`) and data-area merges (`merges`), but not cell styles, header styles or layout features such as width, freeze and filter (a console warning is printed). Keep exports under 50k rows when you need full styling. See [Worker & streaming](/packages/excel-exporter/guide/06-worker-stream).

### Is my data uploaded anywhere?

No. All processing happens in the browser or the Node process; business data never leaves the machine.

### Date columns render as long text, not dates

Without a `format`, `Date` values are written as plain text Excel does not recognize as a date — an ISO string (e.g. `2026-07-01T00:00:00.000Z`) on every path (main / worker / stream). Declare `format: { type: "date" }` (or `datetime`) on date columns: the Workbook path then stores the Excel date serial and auto-injects the matching `numFormat`, so the cell becomes a real date.
