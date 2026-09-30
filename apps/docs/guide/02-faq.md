# FAQ

The questions that come up first across the ecosystem, grouped by package. Each package's own docs go deeper — the [Ecosystem](/guide/) page links to them all.

## General

### Is my data uploaded anywhere?

No. All processing happens in the browser or the Node process; business data never leaves the machine. That holds for every package — export, preview and the engine layer alike.

## Excel export (`@marcusok/excel-exporter`)

### WASM 404 in the browser

With the default (zero-config) resolution this should not happen on Vite / webpack 5 — the bundler emits the shipped `modern-xlsx.wasm` as a hashed asset automatically. If you see a 404, you are likely overriding the URL (`configureWasm({ wasmUrl })` pointing at a wrong path) or using a bundler without `new URL(asset, import.meta.url)` support; point `configureWasm` at a URL your site actually serves, or copy the file out of this package's `dist/` into a static directory.

### WASM loads in the build but not under `vite dev`

Vite's dev server pre-bundles dependencies into `node_modules/.vite/deps/`, where `import.meta.url` no longer points at the package's `dist/`, so the default WASM URL resolves to a path that does not exist. The failure is quiet — the export still succeeds, but styles, widths, freeze panes and filters are stripped (the style-less stream fallback). Exclude the package from pre-bundling and restart the dev server:

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ["@marcusok/excel-exporter"] },
});
```

`vite build` is unaffected. Full diagnosis, including the `expected magic word` compile error and the `WASM load previously failed` diagnostic trap on later attempts: [pre-bundling caveat](/packages/excel-exporter/guide/02-installation#vite-dev-server-pre-bundling-caveat).

### Worker mode degrades to the main thread

The worker asset (`export.worker.js`) resolves automatically; degradation happens when the Worker route fails (e.g. a misconfigured `workerUrl` override 404s). The export then **retries on the main thread** (modern-xlsx keeps styles; the fast stream needs no WASM at all) — the style-less stream fallback is only the last resort when the main-thread retry also fails. Check for `[excel-exporter]` console warnings to find the reason.

### `result.error` is set although `success` is true

The export degraded to the style-less fast stream (WASM failed or is unsupported). Styles are stripped; headers and merges are preserved. The reason is in `result.error.message` and in the `[excel-exporter]` console warnings — usually a 404 wasm URL. See [fallback](/packages/excel-exporter/guide/08-fallback).

### Exporting 100k rows is very slow (>15s)

You are most likely on the `main` + `Workbook.toBuffer()` path, which has a cliff beyond ~55k rows. Keep `mode: "auto"` (~0.8s at 100k rows), or set `mode: "stream"` / `mode: "worker"` explicitly. See [auto mode](/packages/excel-exporter/guide/03-auto-mode).

### Styles do not apply in stream mode

The stream path supports multi-row headers (`children`) and data-area merges (`merges`), but not cell styles, header styles or layout features such as width, freeze and filter (a console warning is printed). Keep exports under 50k rows when you need full styling. See [Worker & streaming](/packages/excel-exporter/guide/06-worker-stream).

### Date columns render as long text, not dates

Without a `format`, `Date` values are written as plain text Excel does not recognize as a date — an ISO string (e.g. `2026-07-01T00:00:00.000Z`) on every path (main / worker / stream). Declare `format: { type: "date" }` (or `datetime`) on date columns: the Workbook path then stores the Excel date serial and auto-injects the matching `numFormat`, so the cell becomes a real date.

## Excel preview (`@marcusok/excel-preview`)

### An encrypted workbook reports `PASSWORD_PROTECTED`

The file is encrypted (Agile AES-256) and no password was supplied. Pass one — `createPreview(el, { source: file, password })`. Note that a missing or wrong password always arrives through `onError` with code `PASSWORD_PROTECTED`; `createPreview` itself never throws, so handle errors in the callback rather than in a `try/catch`. See [Basic Usage](/packages/excel-preview/examples/01-basic).

### My `.xls` / `.ods` file is rejected outright

That is deliberate in v1, and each case gets its own code: legacy `.xls` (BIFF8) is `LEGACY_FORMAT` (re-save as `.xlsx`), `.ods` and other ZIP-based suites are `CORRUPT` (a ZIP archive without an `xl/workbook.xml`), and SpreadsheetML / HTML-table exports are `UNSUPPORTED`. UTF-16 CSV lands in `UNSUPPORTED` as well — re-save it as UTF-8. The full table is in [Scope & Limits](/packages/excel-preview/guide/04-limits).

### Does it read a Chinese `.csv` exported from Excel?

Yes. UTF-8 (with or without BOM) and GB18030 are detected automatically, and the delimiter is sniffed across `,`, `;`, tab and `|`. UTF-16 CSV (`FF FE` / `FE FF` BOM) is the exception — it is detected as binary and rejected with `UNSUPPORTED`, so re-save it as UTF-8 first.

### `setSheet()` right after `createPreview()` does nothing

That is the documented timing, not a bug: `createPreview` boots asynchronously, so `setSheet()` is a silent no-op and `getSheetNames()` returns `[]` until the parse resolves. Call them from `onParsed`, from your own UI, or let the tab bar drive them. `destroy()` likewise does **not** cancel an in-flight parse — it only suppresses the callbacks that parse would have fired. See [Quick Start](/packages/excel-preview/guide/01-quick-start).

### A formula cell renders empty

Formula cells render their **cached** value — the same convention as SheetJS and exceljs; the preview never evaluates formulas. A file written by a non-Excel writer without a cached `<v>` therefore shows an empty cell; opening and re-saving it in Excel populates the cache. A formula with a cached result does render (a numeric cache arrives as `type: "number"`, a string cache as `type: "string"`).

### Does previewing a huge file freeze the page?

No. Parsing runs in a Web Worker (100k × 10 cells measured at ~1.5s) and the renderer mounts viewport cells only, so the tab stays responsive regardless of the file's size. Two real limits remain: the whole file is read into memory inside the worker (multi-hundred-MB files hit memory limits before anything else), and the grid is one tall container element, so Firefox / Safari stop scrolling past ~890k rows (~1.67M in Chrome). Details in [Scope & Limits](/packages/excel-preview/guide/04-limits).

## Progress overlay (`@marcusok/progress-overlay`)

### The overlay never appears for a task that finishes quickly

That is `delayMs`, which defaults to 200ms: a task that finishes sooner never shows the overlay at all, instead of flashing it for a frame. Set `delayMs: 0` to mount immediately.

### It never appears during a long synchronous span

With the default `delayMs: 200`, a long blocking span starves the reveal — the delay expires while the thread is busy, so the overlay never paints. Mount it up front and yield a frame before the blocking call:

```ts
const overlay = showProgressOverlay({ delayMs: 0 });
try {
  await nextPaint(); // let the browser paint the overlay first
  heavySyncWork();
} finally {
  overlay.close();
}
```

`nextPaint()` is exported from the same package. See [Blocking the main thread](/packages/progress-overlay/guide/01-usage).

### Two tasks run at once — whose texts are shown?

Concurrent tasks share one reference-counted DOM node, and the last writer wins: the later caller takes over the display as soon as it calls `showProgressOverlay`, and the node is removed when the last task closes. If an early-quitting caller's pending reveal timer fires after it closed, the texts shown belong to the most recent still-active caller.

### Does it work in Node / SSR?

Yes. Without a `document` the call returns a no-op handle, so the same call site works in both environments — no branching, no `typeof window` check.

## Engine layer (`@marcusok/xlsx-core`)

### Do I have to install it?

No. `@marcusok/xlsx-core` is installed for you as a dependency of both document packages, and both of them **re-export** its `configureWasm` / `getWasmLoader` — so even self-hosting the WASM binary does not require adding it to your own manifest. Install it explicitly only when you build on the engine directly (your own reader, writer or renderer).

### Does `modern-xlsx`'s `engines.node >= 24` affect my app?

No. The engine runtime is bundled **into** `xlsx-core`'s `dist/` at build time, so consumers pull in zero external runtime dependencies and never see that range. `xlsx-core` itself requires Node `>= 22`, and browsers need WebAssembly support. The pinned `modern-xlsx` entry in its manifest exists purely so TypeScript consumers can resolve the re-exported types.

### Two copies of the package in one bundle

Then you have two engines and two loader states — the single-instance guarantee is a module-instance guarantee. Both business packages pin an exact `@marcusok/xlsx-core` version in their published manifests, so one copy is the normal case; a hand-pinned engine in your own manifest is what creates the conflict. Check what a package pins with `npm view @marcusok/excel-exporter dependencies`, and see [Package Relationships & Selection](/guide/03-package-relationships) for the version rules.
