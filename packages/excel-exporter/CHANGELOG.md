# @marcusok/excel-exporter

## 2.6.2

### Patch Changes

- 9ad42b5: Fix the WASM loader getting stuck in its error state after a late success: `initWasm` keeps a single module-level in-flight promise, so on a slow network the underlying load could still succeed after all retried attempts had timed out — the engine was then initialized, yet every later export kept failing with "WASM load previously failed" until `configureWasm()` was called. A late success now flips the loader back to ready.

## 2.6.1

### Patch Changes

- 13e4e3b: Fix the progress overlay staying invisible when a new export starts during the previous overlay's fade-out: the reused node kept `opacity: 0` (from the cancelled teardown) for the whole run while still blocking page interaction. The overlay now fades back in and resets to its initial indeterminate state on reuse.

  Also:

  - `configureWasm({ workerUrl })` now prints a warning when the worker URL changes: the shared Worker reads its script URL once at creation, so a later change only reaches a worker created after `terminateWorker()` (the `wasmUrl` caveat already warned; `workerUrl` did not).
  - The reserved `__index__` prop check is now alias-aware on every path (conflict detection, Workbook and Fast stream builders): a hand-written legacy `key: "__index__"` column is rejected when `indexColumn` is enabled and driven by the row number on the low-level builders, instead of being silently read from `data`.

## 2.6.0

### Minor Changes

- e75e1a2: New optional subpath `@marcusok/excel-exporter/overlay`: a full-screen progress overlay for exports.

  ```ts
  import { exportExcelWithOverlay } from "@marcusok/excel-exporter/overlay";

  const result = await exportExcelWithOverlay(
    { filename: "report", sheets: [...] },
    { delayMs: 200 },
  );
  ```

  - **`exportExcelWithOverlay(options, overlayOptions?)`** wraps `exportExcel`, showing a blocking overlay while the export runs and removing it when the promise settles — on success and on failure alike. It **appends to** your existing `onProgress` / `onPhase` callbacks rather than replacing them, so any UI already wired to those keeps working.
  - **`showExportOverlay(overlayOptions?)`** returns a handle (`handleProgress` / `handlePhase` / `close`) for `exportTable`, `exportEcharts`, and custom flows.
  - Options: `delayMs` (default 200 — a faster export never shows the overlay at all), `minVisibleMs` (default 300), `fadeOutMs`, `zIndex`, `container`, `blockInteraction` (default true), `theme` (`auto` / `light` / `dark`), and `text` label overrides.
  - The overlay is **indeterminate until real progress arrives**: only the Fast stream routes emit `onProgress` values between 0 and 1. Workbook routes (browser < 20,000 rows, Worker 20,000–49,999 rows) show the animated sweep for their whole run. The trailing `onProgress(1)` never closes the overlay and never fabricates a completed bar on a route that had no real progress.
  - `delayMs` also decouples the overlay from main-thread blocking: on a route whose build blocks the thread, a reveal that could not fire before the block **never happens** rather than appearing after the export finished. Use `delayMs: 0` to mount before the export starts and let a two-frame yield paint it first.
  - No `document` (Node/SSR) returns a no-op handle, so the same call site works in both environments. No new runtime dependencies.

  The main entry is untouched — `src/index.ts` and `src/types.ts` are unchanged, and the overlay is a separate build entry, so callers who don't import it ship exactly the same bytes as before.

  Also in this release: a new `Progress Overlay` guide page (English and Chinese), an overlay toggle in the play and docs-site demos, and `overlay-state.test.ts` / `overlay.test.ts` (35 new tests; `happy-dom` added as a test-only devDependency).

## 2.5.0

### Minor Changes

- 9438dff: Fixes and API additions from a full-repo audit:

  - **`exportTable()` now forwards `dataStyle` / `indexColumn`.** Both fields were declared on `TableExportOptions` but dropped by `tableExportToOptions`, so `exportTable({ indexColumn: true })` silently produced no index column (and `dataStyle` silently did nothing) while TypeScript reported no error. The 2.3.0 changelog already claimed the opposite — this release makes it true, with a regression test that goes through `exportTable` itself rather than only `tableToSheet`.
  - **`indexColumn.width` is now validated** like any column width (finite, non-negative). It previously bypassed pre-flight validation because the index column is injected _after_ validation runs: `indexColumn: { width: NaN }` reached the engine, failed with a cryptic serde error and degraded the whole export to the style-less stream, while the same input on the ≥ 50,000-row route silently ignored the width — the exact cross-threshold split the `col.width` check exists to prevent.
  - **New exports: `applyIndexColumn(sheet)` and `INDEX_PROP`.** `WorkbookBuilder.addSheet()` and `exportAsStream()` only understand an already-expanded `__index__` column — they never expanded `SheetConfig.indexColumn` themselves, and there was no public way to do it, so the field was silently ignored on those two entry points. `applyIndexColumn` performs the exact expansion `exportExcel` applies internally.
  - **`FormatSpec` pattern documentation corrected**: the stream path lower-cases the whole pattern and does not interpret quoted literals, so `yyyy"年"M"月"d"日"` renders as `2026"年"m"月"d"日"` above the threshold, where the Workbook path renders `2026年7月1日`. The docs said "verbatim", which was inaccurate for letter case. No behavior change.
  - Documentation-only corrections across the package README and the docs site: preset count (7 → 8), worker-route degradation wording (Worker + stream fails terminally with `success: false` rather than degrading a third time), auto-filter range (last header row plus all data rows), ECharts default header list and `layout` scope, header-row-count formula (`1 + max depth`), and `indexColumn` behavior on the low-level entry points.

## 2.4.0

### Minor Changes

- 757f55a: Refine the default look of styled exports:

  - **Default alignment — every cell is centered.** A base style (`BaseCellStyle`, horizontal + vertical `center`; also exported from the package) is layered underneath all cells on the Workbook path, headers included, so a table no longer mixes Excel's native alignment (text left, numbers right) column by column. Explicit `alignment` values still win field by field — declaring only `horizontal` inherits the base's `vertical` — and either `dataStyle` or `headerStyle` overrides the base. To restore Excel's native alignment, declare both axes explicitly (e.g. `dataStyle: { alignment: { horizontal: "left", vertical: "bottom" } }`).
  - **Index column header text.** `SheetConfig.indexColumn` now defaults its `label` to `"序号"` instead of `"No."`. Pass an explicit `label` to display something else.
  - **Darker preset borders.** `StylePresets.dataRow` and `StylePresets.bordered` now use `BFBFBF` instead of `D0D0D0`, so their thin borders read clearly rather than washing out.

## 2.3.0

### Minor Changes

- 6b21f86: Add sheet-level table styling and a built-in index column:

  - `SheetConfig.dataStyle`: base `CellStyle` applied to every data cell. A column's `style` deep-merges over it field by field — a table-wide border survives a column that only sets `numFormat`, and vice versa. Together with the new `StylePresets.bordered` preset (thin `D0D0D0` box on all four sides), a fully bordered table is now `dataStyle: StylePresets.bordered`: one field, no per-column repetition.
  - `SheetConfig.indexColumn` (`true`, or `{ label, width, start, style, headerStyle }`): injects a leading row-number column. Values are generated from the row number (never read from `data`), existing `merges` shift one column right automatically, and the feature works on every export path — workbook, worker, and the style-less stream alike. `exportTable` accepts both new fields too.

## 2.2.0

### Minor Changes

- cee8203: Align column field naming with Element Plus: `prop` (data-row field) and `label` (header text) are now the canonical `ColumnConfig` fields. The pre-2.2 names `key` / `header` keep working as deprecated aliases — `prop` / `label` win when both are present, so existing code exports unchanged. `exportTable()` resolves names as `prop ?? key ?? dataIndex` and `label ?? header ?? title`. Columns providing neither `label` nor `header` are now rejected up front with a clear error.

## 2.1.6

### Patch Changes

- 8ac0e47: Progress-contract and documentation fixes from a full-repo audit:

  - The stream path's final 1,000-row progress checkpoint is now skipped when it would land exactly on the last row: whole-thousand row counts previously emitted a duplicate `onProgress(1)` immediately before the terminal 1 that the `onProgress` contract promises fires exactly once. Regression tests added at both the engine and the entry level.
  - `FormatSpec` pattern-token docs now state precisely that superset tokens are only partially verbatim on the stream path: `"mmm"` parses its `mm` prefix and emits a stray `m` (`"mmm"` -> `"09m"`), while the Workbook path renders the month abbreviation via numFormat.
  - The `configureWasm` runtime warning and JSDoc now name the import location of `terminateWorker()` (`@marcusok/excel-exporter/worker-utils` subpath — it is not exported from the main entry, so the warning previously pointed at a function callers could not resolve).
  - Internal cleanup, no behavior change: `buildWorksheetXml` parameters are now required (the optional declarations contradicted the non-null-asserted implementation and the single call site), and the duplicated `columnName` helper was consolidated into the shared `column-tree` export.

## 2.1.5

### Patch Changes

- 885a8a3: Fix cross-path input handling and degradation-chain edge cases found in a full audit:

  - Node/SSR stream-route build failures now fail directly instead of re-running the identical fast-stream fallback on the same input (parity with the worker-chain guard) — no more doubled build time and duplicated warnings on a doomed retry.
  - Pre-flight validation now covers numeric fields: `width` (finite, non-negative; `0` stays legal and hides the column), `freezeRows` (non-negative integer), `format.decimals` (integer 0–100), and `padding.length` (integer 0–10000). Previously a NaN/negative `width` or fractional `freezeRows` failed the Workbook path with an opaque engine error and degraded the whole export to the style-less stream, while the same input was silently ignored (or threw a `toFixed` RangeError) on the stream path — the same config could succeed below 50k rows and fail above them.
  - `exportTable` rejects circular `children` with a clear error instead of a stack overflow (the adapter's recursive mapping ran before the generic cycle check).
  - `configureWasm`/`updateOptions` compare `wasmUrl` by string, so re-passing an equal `URL` object no longer resets an already-configured loader; the idempotence warning now also covers the "initial fetch still in flight" case (an in-flight `initWasm` cannot be redirected — documented in the caveat).
  - `WasmLoader` state writes are centralized with the promise-identity check: a superseded load (URL changed during a retry backoff) can no longer overwrite the successor's ready state, so `getWasmLoader().isReady` no longer flips back to false.
  - Non-`Error` rejections keep their original message in the WASM load error and in worker failure responses (previously `undefined` / "worker unknown error").
  - Build: the wasm copy step moved from the `build` script into tsup's `onSuccess` hook, so `tsup --watch` (`pnpm dev`) rebuilds restore `dist/modern-xlsx.wasm` after the clean step wipes it — dev sessions no longer break the Node auto-init path and integration tests.
  - Types: corrected `onPhase`/`ExportResult.duration` docs (worker-route duration includes main-thread serialization and the worker round-trip, not in-worker time only) and `autoFilter` docs (the filter range spans the last header row plus all data rows).
  - Tests: new direct coverage for the worker message protocol (success path, init idempotency via the string-normalized URL key, byte transfer, error stringification), which previously had only the timeout branch tested.

## 2.1.4

### Patch Changes

- cb3f5f5: Robustness fixes from a code audit. Invalid Date values no longer crash every export path with a raw `RangeError` (they render as `"Invalid Date"` strings, matching `String(new Date(NaN))`). Enum FormatSpec lookups no longer walk the prototype chain: values like `"constructor"` / `"__proto__"` that the map does not define now hit the fallback instead of writing `function Object() {...}` text into the cell. Object-form ECharts scatter datums (`{ value: [x, y] }`) are now exported as X/Y pairs exactly like bare `[x, y]` pairs instead of being silently stringified as `"[1,2]"`. A missing or empty `filename` fails fast with a structured `{ success: false }` error instead of returning `success: true` with no file and a cryptic warning. The browser download trigger releases its object URL even when a sandboxed DOM throws mid-trigger.

## 2.1.3

### Patch Changes

- 43ef12a: Docs and filename edge case:

  - `download.ts`: the `.xlsx` suffix check is now case-insensitive, so a filename like `report.XLSX` no longer becomes `report.XLSX.xlsx`.
  - `types.ts` doc comments: `FormatSpec` now documents that the stream path parses only the `yyyy`/`MM`/`dd`/`HH`/`mm`/`ss` pattern tokens and emits anything else verbatim (the Workbook path renders any valid Excel format code via numFormat); `ExportPhase` now states that the browser worker-stream route reports no `init` phase and that a build attempt failing inside the worker reports no `build` phase (main-thread attempts report in a `finally`); the `onPhase` note on `ExportResult.duration` now reflects that the worker route measures in-worker time only.

## 2.1.2

### Patch Changes

- 6dc02e3: Docs accuracy pass (no behavior change):

  - README: `onPhase` now documented as per-attempt (a degradation chain reports one `build` phase per attempt, matching the `ExportPhase` type contract); the WASM-needs table entry now includes auto/Node >= 50,000 rows alongside explicit `mode: "stream"` as WASM-free routes.
  - `types.ts`: the `FormatSpec` / `ColumnConfig.format` comments now state precisely where function-form formats run (all main-thread routes, including Node's main-thread stream and degradation retries) and where they are stripped with a warning (browser worker routes).

## 2.1.1

### Patch Changes

- 9d8ba7f: Fix several robustness details across export routes:

  - Validate that every sheet entry is a non-null object up front (`each sheet must be an object`), so structurally malformed input no longer falls into the export routes with a bare TypeError.
  - Isolate the browser download trigger from the fallback chain: if `triggerDownload` throws (e.g. a mocked/blocked DOM), the export no longer rebuilds via main-thread retry or stream fallback — the Blob is still returned and a warning is logged.
  - Render `null`/`undefined` data rows as empty rows in both workbook and stream paths instead of throwing, including number/date `FormatSpec` columns.
  - Skip empty cells in the fast-xlsx stream writer (no empty-string interning, no empty `<c>` elements), slightly reducing output size.
  - Skip empty `border: {}` style objects instead of emitting a no-op border definition.
  - `echartsToSheet`: throw a clear error when `categoryHeader` collides with the internal `__series_N` keys in wide layout.
  - Add the missing `phase`/`duration` fields to the `WorkerResponse` type (the worker already sent them at runtime).

## 2.1.0

### Minor Changes

- ac935bc: Configurable worker timeout, input-validation hardening, and degradation-chain fixes:

  - **feat: `configureWasm({ workerTimeoutMs })`** — the worker export timeout (previously a fixed 120s) is now configurable. A timed-out export still terminates the shared worker and rejects its sibling requests, so raise it only for legitimately huge exports.
  - **fix: empty `sheets: []` no longer produces a corrupt file reported as success.** Previously the Workbook build threw, the export degraded to the fast stream and resolved `success: true` with a zero-sheet workbook Excel flags as corrupt (ECMA-376 requires at least one sheet). `exportExcel` pre-flight validation and `exportAsStream` now both reject with a clear `at least one sheet is required` error.
  - **fix: a synchronous `postMessage` failure no longer kills the shared worker 120s later.** Un-cloneable payloads (e.g. Symbol/function values in row data) throw DataCloneError synchronously; the request previously stayed in the pending map with its timeout timer live, so the timer eventually fired and terminated the healthy shared worker, also rejecting every sibling request. The pending entry and timer are now cleaned up immediately.
  - **fix: worker stream-route failures stop after one main-thread retry.** On the >= 50,000-row stream route the main-thread retry already runs the identical fast stream on identical input, so the previous third attempt via the terminal fallback failed deterministically (and logged a misleading "falling back" warning). The chain now fails directly with the combined reason.
  - **fix: reusing the same `ColumnConfig` object in two places (a diamond, not a cycle) now throws.** It previously passed the path-based cycle check and was walked twice, silently emitting duplicate data columns.
  - **fix: sheet names beginning or ending with an apostrophe are rejected** (Excel refuses such names), matching the other ECMA-376 name validations.
  - **fix: custom date/time patterns interpret `mm` as minutes when followed by `ss`** (Excel's numFormat convention), so e.g. `mm:ss` renders `20:30` on the stream path instead of month:seconds — consistent with the Workbook path, which passes the pattern to Excel as a numFormat.
  - **docs:** the docs-site package home pages (en + zh) still described the pre-2.0 dependency model (install `modern-xlsx` explicitly, peerDependencies incl. an optional `xlsx` fallback, mandatory asset deployment) — updated to the zero-runtime-dependency, zero-configuration reality; option tables document `workerTimeoutMs`; sheet-name constraints and the stream-route retry semantics are documented (en + zh).
  - **tests:** 108 → 116 (empty-sheets pre-flight on both paths, postMessage-throw cleanup with no delayed worker kill, single stream retry on the stream route, diamond column reuse, apostrophe sheet names, custom worker timeout, `mm:ss` minute disambiguation).

## 2.0.0

### Major Changes

- 2bd9bd9: Zero-configuration assets and a dependency-free install — **breaking** (public `ExportResult.engine` union narrows, the SheetJS fallback is removed, the `xlsx` optional peerDependency is gone):

  - **feat: assets locate themselves — no `main.ts` wiring, no bundler config.** `modern-xlsx.wasm` and `export.worker.js` default to `new URL(<file>, import.meta.url)` next to the package entry. Verified end-to-end on Vite 8: dependency pre-bundling (dev) rewrites both URLs into `node_modules/.vite/deps`-served assets, and production builds emit them as hashed assets (`vite:asset-import-meta-url`, `node.js:29049/28618` in vite 8.2.0); also verified against a packed tarball in a fresh Vite project with an intentionally empty `vite.config.ts`, and in Node against a real `pnpm install`ed copy. `configureWasm` stays as the optional escape hatch (self-hosted CDN copies, Service Worker, bundlers without asset-URL support); the previous `?url` + `configureWasm` setup keeps working.
  - **feat: zero runtime dependencies.** modern-xlsx and fflate are bundled into the package at build time (browser resolution, so fflate's Node `createRequire` shim never leaks into consumer browser builds — that shim hard-fails Vite 5 builds), and the wasm binary ships under this package's own `exports` map. Consumers install one package; upstream `engines.node>=24` declarations no longer affect their installs, and the README's `engine-strict` guidance is retired. The published `.d.ts` no longer imports from `modern-xlsx` (`BorderStyle` is inlined; dts `resolve: true`).
  - **feat: `export.worker.js` is now a single self-contained ESM file** (splitting off, zero imports). This also **fixes a production bug**: the old chunked worker referenced `./chunk-*.js` siblings that `?url` asset emission never copies, so worker mode silently 404'd and degraded to the main thread in every production build (observed in `packages/play/dist`). The default worker URL is deliberately the hoisted `const url = new URL(...); new Worker(url)` form — the inline form makes bundlers re-bundle the worker as an entry, which hard-fails Vite 5 (VitePress) builds (`Invalid value "iife" ... code-splitting builds`).
  - **feat!: SheetJS fallback removed; terminal degradation is now the pure-JS fast stream.** Degradation chain: worker failure → main-thread retry (styles preserved) → style-less fast stream (headers/merges preserved, no WASM, no network). Same degradation surface as SheetJS with none of its costs: no optional peerDependency, no runtime CDN load, no unmaintained-dependency CVEs. **Breaking:** `ExportResult.engine` narrows to `"modern-xlsx"` (degraded exports report `mode: "stream"` + `result.error` with `success: true` — monitor via `result.error` instead of `engine === "sheetjs"`); the optional `xlsx` peerDependency is removed; Node consumers pre-warming via an externally installed `modern-xlsx`'s `initWasmSync` no longer pre-warm the bundled engine — use `await getWasmLoader().ensureLoaded()` at startup instead.
  - **fix(wasm-loader):** Node auto-init now resolves this package's own `dist/modern-xlsx.wasm` (import.meta.url-relative, with a `../dist` fallback for src-mode) instead of `require.resolve("modern-xlsx")` — which would have broken now that modern-xlsx is bundled rather than installed. Zero-config Node export verified against a packed tarball install (`engine: "modern-xlsx"`, `mode: "main"`, no error).
  - **build:** an esbuild plugin rewrites modern-xlsx's dead `new URL('modern_xlsx_wasm_bg.wasm', import.meta.url)` default (a filename this package never ships) to a computed `../dist` specifier — invisible to bundler static analysis, so consumer builds no longer see `doesn't exist at build time` warnings for it; if the branch ever runs it still resolves to the shipped binary (sha256-identical). A second plugin drops modern-xlsx's dead `await import("node:fs/promises")` branches (`toFile`/`readFile` — Node-only APIs this package never exports), replacing them with a rejected promise so consumer browser builds no longer see `Module "fs/promises" has been externalized` warnings either (the loader's own live `import("node:fs")` for Node auto-init stays, matching pre-2.0 baseline).
  - **tests:** 108 passing (was 108 at 1.3.0 with different composition): `fallback.test.ts` replaced by `stream-fallback.test.ts` (drives the new terminal degradation through `exportExcel` with WebAssembly stubbed out — degradation markers, progress contract, grouped headers + merges, multi-sheet rowCount, feature-drop warnings); auto-init unit tests cover the new resolution order (dist-next-to-entry primary, `../dist` src fallback, URL-not-undefined guard) and the default-URL `initWasm` path; routing/input-validation suites updated for `engine: "modern-xlsx"` + `mode: "stream"` on degraded routes.
  - **docs:** package README and the docs site (en + zh) restructured around zero-config: installation is one command, the asset-resolution section documents the default order and the `configureWasm` escape hatch, the fallback/node-ssr/faq pages describe the stream degradation, and two stale instructions are fixed (docs no longer ask consumers to install `modern-xlsx` explicitly, and the Node page no longer claims manual `initWasmSync` is required — both predated 1.3.0). The docs site's live `ExportDemo` drops its `configureWasm` + public-asset wiring (verified: VitePress/Vite 5 builds emit both assets hashed — `dist/assets/modern-xlsx-*.wasm`, `export.worker-*.js`), the `xlsx` devDep is removed from the docs app, and the internal design doc carries a "2.0 status" banner pointing to the current behavior. The play demo drops its three lines of `configureWasm` boilerplate; play's vite config gains a src-alias transform mapping the default asset literals to the package's dist copies (src mode has no dist next to the entry).

## 1.3.0

### Minor Changes

- d470ba9: Simplify asset integration and Node setup:

  - feat: `modern-xlsx.wasm` is re-published under this package's own `exports` map (`@marcusok/excel-exporter/dist/modern-xlsx.wasm`, forwarded at build time). Vite consumers now integrate assets with two `?url` imports + one `configureWasm` call — no copy plugin, no `public/` setup. The previous ~30-line copy plugin remains documented as the fallback for bundlers without asset-URL imports (its copy source also collapses to this package's `dist/`).
  - feat: Node/SSR zero-configuration — when no `wasmUrl` is configured, the engine locates `modern-xlsx.wasm` through `node_modules` via `createRequire` (pnpm-symlink-safe, verified against a published-tarball install) and initializes it synchronously on first use (~20ms measured: ~4ms read + ~15ms compile). The six-line `initWasmSync(readFileSync(...))` entry boilerplate is no longer required; the explicit form stays supported for init-timing control, and any auto-init failure silently falls back to the previous `initWasm` → SheetJS degradation chain.
  - refactor: `modern-xlsx` moved from `peerDependencies` to `dependencies` — consumers install one package, and the wasm binary always ships with the matching JS glue (its own `exports` map omits wasm subpaths, which is why the forwarding exists). `xlsx` (SheetJS) remains an optional peerDep. Engines note unchanged: modern-xlsx declares `node>=24`; Node 22 works and CI runs there.
  - docs: README + docs-site installation/getting-started/FAQ/Node-SSR pages restructured around the layered story — `wasmUrl` is the only always-needed asset (explicit `mode: "stream"` excepted), `workerUrl` is needed only for Worker routes (auto >= 20k rows).
  - tests: 100 → 107 (CI runs 103): Node auto-init unit tests (mock-missing-initWasmSync fallback, initWasmSync-throw degradation, configured-URL opt-out, browser-scope opt-out), a real-wasm integration file (zero-config load + real `engine: "modern-xlsx"` export), and a mock-compat fallback case. Package vitest now runs files serially (`fileParallelism: false`): the real-wasm files' CPU spikes previously pushed the timing-based performance baselines past their SLA thresholds under parallel workers (observed 204-273ms vs ~120ms isolated).

## 1.2.3

### Patch Changes

- e62b829: - fix: a timed-out worker export (fixed 120s) no longer leaves the worker cached and running — the worker is terminated and dropped so the next export creates a fresh one, and sibling requests dispatched to the same worker are rejected immediately (their callers degrade to the main-thread retry) instead of queueing behind a wedged worker.
  - fix: `format: { type: "number" }` columns now render `null`/`undefined` values as empty cells on every path. Previously `Number(null) === 0` silently turned missing values into `0` while `undefined` became `""` (asymmetric, and `0` is meaningful in financial data).
  - fix: the SheetJS fallback warning now lists configured layout features it drops (`width` / `freezeRows` / `autoFilter`) alongside "styles stripped", matching the stream path's per-feature warnings instead of degrading them silently.
  - fix: `configureWasm()` with a new `wasmUrl` after a successful load now warns that modern-xlsx's idempotent `initWasm` keeps the already-loaded module on that thread (the previous JSDoc/README claim that the next `ensureLoaded` "re-initializes from the new URL" was not what the real dependency does — verified against modern-xlsx 1.2.0's module-level `initialized` guard). Tests updated to describe the loader's actual state-machine contract.
  - perf: the Workbook path deduplicates identical cell styles — the same `CellStyle` used by N header cells or columns now shares one style index instead of appending N duplicate font/fill/xf records (modern-xlsx's `StyleBuilder.build` never dedupes; 50 styled header cells previously produced 50 identical records).
  - docs: `ColumnConfig.format` / `ExportResult.error` doc comments now state the `thousands` cross-path difference (Workbook renders `#,##0` via numFormat; stream/SheetJS keep the cell a number without separators), the null/undefined-as-empty-cell rule, and that `error` is also set on `success: true` fallback results (degradation reason).
  - chore: the published tarball now includes `CHANGELOG.md`.
  - tests: 94 → 100 (CI runs 96): worker timeout termination + sibling rejection, style dedup round-trip, fallback dropped-features warning, wasmUrl-change warning, number-spec null/undefined rendering.

## 1.2.2

### Patch Changes

- 59ae680: - fix: the Workbook path now normalizes non-primitive cell values exactly like the stream and SheetJS paths — plain objects are written as JSON strings, `Date`s as ISO strings, bigints as decimal strings — so a dataset crossing the 50k-row threshold (or exported before/after a degradation) keeps identical cell content. Previously objects landed as `"[object Object]"` and dates as locale-dependent long text on the Workbook path only.
  - fix: `toStr` now handles `symbol`/`function` values explicitly (`JSON.stringify` returns `undefined` for them), so such cells receive a visible string instead of `undefined`.
  - fix: the worker no longer re-runs `initWasm` and re-reports the `init` phase on every export: a URL-typed `wasmUrl` arrives via structured clone as a fresh object each message, so the old reference comparison never matched. Initialization is now tracked by the URL's string form plus an explicit ready flag. Note `initWasm` is idempotent — a `wasmUrl` change after the first successful init does not take effect (the stale comment claiming otherwise is corrected).
  - fix: when the browser Worker route fails (missing/404 `workerUrl`, WASM init error in the worker, timeout), `exportExcel` now retries on the main thread with modern-xlsx first (styles preserved; the ≥50k-row tier uses the WASM-free fast stream) and only degrades to the style-less SheetJS fallback if that retry also fails. The `onProgress` 0→1 contract (each exactly once) is unchanged.
  - tests: 91 → 94 (CI runs 90): cross-path value-normalization regression (Workbook vs stream read back cell-by-cell), worker-failure → main-thread retry (workbook and stream tiers), and the full worker → main-thread → SheetJS double-failure chain.

## 1.2.1

### Patch Changes

- b26e83f: - fix: reject a sheet with an empty `columns` array up front with a clear error (`at least one column`) on every path, instead of crashing the Workbook autoFilter layout with a cryptic `TypeError` (`encodeCellRef(0, -1)` → `"@1"`) and then silently degrading to the SheetJS fallback.
  - fix: `echartsToSheet` long/item layouts now reject duplicated header texts (e.g. `seriesHeader` equal to `valueHeader`) with a clear error — header texts double as row keys there, so duplicates previously overwrote each other's column silently.
  - docs: correct the stale test-count figure in the README and document the Chinese default headers of `exportEcharts` (overridable via the `*Header` options).

## 1.2.0

### Minor Changes

- 156584b: - fix: the Workbook path (default route under 50k rows) now writes `NaN`/`Infinity` as visible strings instead of illegal `<v>NaN</v>` XML that Excel flags as corrupt, matching the stream and SheetJS paths.
  - fix: `exportExcel` now fails invalid input (bad merges, duplicate/invalid sheet names, missing column keys) immediately via a pre-flight check, instead of first attempting a pointless SheetJS fallback — same error messages as before, no misleading "Falling back" warning.
  - The SheetJS fallback result's `error` message now carries the degradation reason programmatically (e.g. `workerUrl not configured`).
  - feat: export the `LoaderOptions` / `LoadState` types from the package entry (documented but previously missing).
  - docs: document the SheetJS npm CVE situation; installing the optional `xlsx` peer from the official CDN tarball is now recommended.

## 1.1.1

### Patch Changes

- 9d9339f: 修复多路径输入校验与进度契约问题：

  - stream/SheetJS 路径遇 `NaN`/`Infinity` 不再写出非法 `<v>NaN</v>` XML（此前 Excel 会判定文件损坏），改为写入可见字符串 `"NaN"`/`"Infinity"`；workbook 路径行为不变（空单元格）。
  - 新增 `merges` 统一校验（整数、`row`/`col` ≥ 0、`rowspan`/`colspan` ≥ 1、不超出数据区、互不重叠），三条路径一致以 `{ success: false, error }` 失败并指明问题项，替代原先直接生成损坏文件的行为。
  - 重复 sheet 名三条路径统一报错（此前 stream 路径产出损坏文件、workbook 路径意外降级丢样式、SheetJS 路径静默改名）。
  - worker 失败降级 SheetJS 的路径修复 `onProgress(1)` 双发，恢复「trailing 1 恰好一次」契约（types.ts）。
  - stream 模式特性跳过警告改为递归扫描列树，多级表头下嵌套的 `width`/`style`/`headerStyle` 不再被静默丢弃，兑现 README「dropped with a warning」承诺。
  - 新增 17 个测试（输入校验 + worker 进度契约），总计 84 个用例。

## 1.1.0

### Minor Changes

- e92f6d6: feat: 支持多级表头与单元格合并

  - `ColumnConfig` 新增 `children`：列可组成树形结构，生成多行表头。分组表头格自动跨其全部叶子列合并，叶子表头格纵向跨满剩余表头行，无需手工计算合并范围。
  - 多级表头与合并（含表头合并）在 Workbook / Fast stream / SheetJS 兜底三条路径均可用；数据区 `merges` 的相对偏移随表头行数自适应，扁平列配置的输出与旧版逐字节一致。
  - `exportTable` 支持 Ant Design / Element Plus 的 `children` 分组列。

## 1.0.8

### Patch Changes

- 08239dd: 修复 onProgress 兜底契约与 sharedStrings `count` 规范偏差：

  - `onProgress` 的收尾 `1` 此前只在成功路径由 `exportExcel` 上报，三条 SheetJS 兜底路径（WASM 不支持早退、主线程构建失败降级、Worker 失败/抛错降级）均不上报，早退路径连起始 `0` 也不上报，与 `types.ts`「final 1 由 `exportExcel` 恰好上报一次」的契约不符。现 `exportExcel` 在入口统一上报 `0`，兜底调用统一经 `.finally` 收尾 `1`（兜底自身失败也收尾），任何路径下回调序列均为 `0 → … → 1` 各一次，进度 UI 可确定性关闭。
  - fast-xlsx 的 `xl/sharedStrings.xml` 原把 `count` 与 `uniqueCount` 同填去重数；按 ECMA-376，`count` 应为含重复的总字符串引用数。现按实际引用计数填写（Excel 等读取器原本也容忍该偏差，属规范正确性修正）。
  - 清理 `PERF_TIGHT` 残留：性能基准的 `SLACK` 恒等式（两分支同为 1.0）改为直赋 `1.0` 并修正注释；`turbo.json` `globalEnv` 移除无效的 `PERF_TIGHT` 声明。
  - 新增两个回归用例（兜底路径进度收尾、sst count/uniqueCount 规范），测试数 52 → 54（CI 跳过 4 个性能基准后实跑 50）。

## 1.0.7

### Patch Changes

- cd54ef1: README 修正函数形式 `format` 的适用范围表述：

  - 原「`main` 模式额外支持函数形式」与设计决策摘要「Worker/Stream 仅接受 `FormatSpec`」均不完整：Node 的 stream 路径在主线程执行，函数同样生效；会剥离函数的只有浏览器 Worker 路径（含 Worker 内执行的 stream）。
  - 同步修正文档站 FAQ 的日期条目（按 Workbook / stream·SheetJS 兜底路径分述默认文本形态）与基准图 caption（注明 6 列测量口径与在线演示 9 列数据集不可直接对照）。

## 1.0.6

### Patch Changes

- 8b28fab: 修改文档
- 8b28fab: 文档与包元数据措辞修正：

  - `package.json` description 移除 "High-performance" 营销化措辞，改为事实性描述（与 1.0.4 清理 README 同类措辞的决定对齐）。
  - README 修正测试数量为当前实际值（共 52 个用例，CI 跳过 4 个性能基准后实跑 48 个；此前写的 47/43 已漂移）。
  - README 环境要求澄清：`pnpm >= 9` 是本仓库的开发环境要求，不是消费方的安装要求。

## 1.0.5

### Patch Changes

- 480737b: 跨路径一致性修复：

  - `date` / `datetime` 统一按 **UTC 分量**解释：Stream/SheetJS 路径的 pattern 字符串此前取本地分量，与 Workbook 路径 `dateToSerial`（UTC 口径）相反，非 UTC 时区下同一输入在 5 万行阈值两侧（或降级前后）可能相差一天。现两条路径在任何时区输出一致；日期列建议传 ISO 字符串或用 `Date.UTC(...)` 构造（见文档「值格式化」的时区约定）。
  - WASM 加载失败（error 态）后，任意 `configureWasm()` 调用都会清除错误态并在下次导出时按新配置重试（此前仅 `wasmUrl` 变化才会重置，错误信息建议的重试方式实际不可行）。
  - stream 路径不再在内部重复上报最终进度 `1`，由 `exportExcel` 统一收尾（回调序列 0 → 分段 → 1 各一次）。
  - 新增 `wasm-loader.test.ts` 与日期跨路径一致性回归用例；`types.ts` 补日期输入契约与 `onProgress` 精确语义。

## 1.0.4

### Patch Changes

- 6046e9b: README（npm 页面）文案修正：

  - 移除「高性能」「降级保底」「保证数据可用」等无基准或绝对化的措辞，改用实测口径描述。
  - 性能测量口径说明更新为「6 列混合类型」，移除与当前 Play 演示（9 列）不符的「Play 同款 6 列」表述。

## 1.0.3

### Patch Changes

- c6e0964: README（npm 页面）更新：

  - 顶部新增在线文档链接：<https://yourbusiness.github.io/marcusok/packages/excel-exporter/>。
  - 修正 peerDependency 安装说明：npm 7+ / pnpm 8+ 起默认会自动安装 peerDependency（原「pnpm 默认不自动安装 peerDep」的说法已过时），显式安装的意义在于锁定版本意图。
  - 修复指向 `docs/excel-export-design.md` 的相对链接在 npm 页面失效的问题（改为 GitHub 绝对链接）。

## 1.0.2

### Patch Changes

- 12d47a4: Fix three robustness issues found in a code review:

  - **wasm-loader race**: calling `configureWasm()` with a new `wasmUrl` while a load was in flight left the loader marked ready with the _old_ URL's WASM (the superseded load clobbered the reset state). The in-flight promise is now captured locally so a superseded load can no longer mark the loader ready/error; the new URL takes effect on the next `ensureLoaded()`.
  - **broken worker reuse**: after a `Worker` `onerror` (e.g. failed script load), the errored instance stayed cached and every later export failed into the style-less SheetJS fallback. The errored worker is now terminated and dropped so the next export creates a fresh one; only requests dispatched to that instance are rejected.
  - **`download` phase in Node**: `onPhase("download", ...)` was reported in Node even though no download can happen there, contradicting the documented `ExportPhase` contract. The phase is now only reported when a browser `document` exists.
  - Stream mode now also warns when data-cell column `style`s are dropped (previously only `headerStyle`/`width`/layout features warned); the console message changed from `stream mode: layout features not supported (...)` to `stream mode: features not supported (...)`.

## 1.0.1

### Patch Changes

- 6194890: 大文件导出切换为自研 Fast stream（fflate minimal OOXML），修复浏览器 Worker 回调克隆失败导致降级 SheetJS 的问题，并将 10 万行导出耗时降至 1000ms 以内。

## 1.0.0

### Major Changes

- 4ad5ee1: update docs & finish beta

## 0.4.0

### Minor Changes

- 0c0fbd5: 性能优化

### Patch Changes

- 0c0fbd5: 大文件导出路径切换为 fflate-based fast-xlsx，10 万行 4 列首次导出从约 1.5s 降至约 600ms，恢复 5 万行 <500ms / 10 万行 <1000ms 硬性指标。

## 0.3.1

### Patch Changes

- fee37db: 添加下载模式

## 0.3.0

### Minor Changes

- bbb89a6: 添加在线文档

### Patch Changes

- ed3e961: Expose `./package.json` in the exports map so consumers (e.g. the docs site) can read the installed version at runtime.

## 0.2.0

### Minor Changes

- cb98c84: feat: `exportExcel` 新增 `onPhase` 阶段耗时回调（`init` / `build` / `download`），
  每个阶段完成时上报实际毫秒数，便于 play 指标面板展示下载链路的分阶段耗时。

### Patch Changes

- f182d80: fix some config error
- cb98c84: 添加了play

## 0.1.3

### Patch Changes

- c474368: 改了一下小配置

## 0.1.2

### Patch Changes

- 31b0cfe: 修复了部分代码问题

## 0.1.1

### Patch Changes

- cefad0e: 修改了一些配置文件
- cefad0e: Tighten package `exports`: add a `default` condition to each entry so resolvers/bundlers that do not understand the `import` condition can still resolve the ESM entry points (the package is ESM-only, `type: "module"`; this is not CommonJS/CJS support), expose `./dist/export.worker.js` as a resolvable subpath, declare `xlsx` as an optional peer dependency, and add `@vite-ignore` to the SheetJS dynamic import so builds do not fail when xlsx is not installed. Also fixed a typo in the WASM loader error message.
