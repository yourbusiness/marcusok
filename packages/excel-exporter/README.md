# @marcusok/excel-exporter · Excel Export Engine

An Excel export library built on [modern-xlsx](https://github.com/ABCrimson/modern-xlsx) (WASM) plus a custom Fast stream writer. It offers a declarative API, automatic mode routing, full cell styling, Web Worker multithreading, fast writes, and a style-less pure-JS stream fallback.

> 📖 **Online docs**: https://yourbusiness.github.io/marcusok/packages/excel-exporter/

## Performance Baseline

Measured locally (real Chrome, 6 mixed-type columns; the Node standalone regression test uses a reduced 4-column set, see `src/__tests__/performance.test.ts`):

| Rows      | auto route  | Measured | Hard requirement |
| --------- | ----------- | -------- | ---------------- |
| 10k rows  | Workbook    | ~120ms   | < 200ms          |
| 50k rows  | Fast stream | ~400ms   | < 500ms          |
| 100k rows | Fast stream | ~780ms   | < 1000ms         |

> The large-file path no longer relies on modern-xlsx's `StreamingXlsxWriter`; it synchronously compresses a minimal OOXML workbook with `fflate`. It finishes the 100k×6-column scenario in ~0.8s and avoids the super-linear cliff of `Workbook.toBuffer()` beyond ~55k rows.

## Installation

```bash
pnpm add @marcusok/excel-exporter
```

That is the entire setup. The package has **zero runtime dependencies** — the engine (modern-xlsx JS glue + fflate) is bundled in at build time, and the 1.9MB `modern-xlsx.wasm` binary ships under this package's own `exports` map (`@marcusok/excel-exporter/dist/modern-xlsx.wasm`), so the binary always matches the JS glue. No engine package to install, no `engine-strict` conflicts from upstream `engines` ranges, no fallback package, nothing to wire in `main.ts`. Bundler config is needed in exactly one case — Vite's dev server (see [How assets resolve](#how-assets-resolve-zero-configuration) below).

## Usage

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "sales-2026",
  sheets: [
    {
      name: "Sales",
      freezeRows: 1,
      autoFilter: true,
      columns: [
        { prop: "orderId", label: "Order ID", width: 18 },
        {
          prop: "amount",
          label: "Amount",
          width: 12,
          style: StylePresets.currency,
        },
        {
          prop: "status",
          label: "Status",
          width: 10,
          format: {
            type: "enum",
            map: { paid: "Paid", pending: "Pending" },
            fallback: "Unknown",
          },
        },
      ],
      data: [{ orderId: "ORD-001", amount: 9999.99, status: "paid" }],
    },
  ],
});
```

Columns follow Element Plus naming (`prop` = data-row field, `label` = header text). The pre-2.2 names `key` / `header` still work as deprecated aliases (`prop` / `label` win when both are present), so existing code keeps exporting unchanged.

## How assets resolve (zero configuration)

Two files ship alongside the code and are located automatically:

- **`modern-xlsx.wasm`** (1.9MB, the style engine) — needed by the styled routes (main / worker + Workbook); the Fast stream routes (explicit `mode: "stream"`, or auto/Node ≥ 50,000 rows) do not use WASM at all.
- **`export.worker.js`** (self-contained, zero imports of its own) — needed only when exports actually enter a Worker (auto mode with ≥ 20,000 rows, or an explicit `mode: "worker"` / `mode: "stream"`).

Both default to `new URL(<file>, import.meta.url)` relative to the package entry:

- **Bundlers** — production builds (Vite, webpack 5) rewrite the expression and emit the file as a hashed asset; nothing to import, copy or configure. **One exception: Vite's dev server.** Dependency pre-bundling (`optimizeDeps`) does not rewrite the expression inside pre-bundled dependencies, so the URL points into `/node_modules/.vite/deps/` where the file does not exist — the HTML fallback returns a page instead, WASM compilation fails, and exports silently degrade to the style-less stream (styles/widths/freeze stripped, `result.error` mentions `Fallback: styles stripped`). Fix it once in `vite.config.ts`, then restart the dev server:
  ```ts
  import { defineConfig } from "vite";

  export default defineConfig({
    optimizeDeps: { exclude: ["@marcusok/excel-exporter"] },
  });
  ```
  Details and alternative wiring (`?url` imports + `configureWasm`): see the [Installation guide](https://yourbusiness.github.io/marcusok/packages/excel-exporter/guide/02-installation).
- **Node** reads the binary from disk next to the installed package and initializes it synchronously (`initWasmSync`) — no fetch, no boilerplate (see [Node Usage](#node-usage)).

`configureWasm` remains as an optional escape hatch for setups where the defaults cannot work — self-hosted copies on a CDN, Service Worker environments, or bundlers without asset-URL support:

```ts
import { configureWasm } from "@marcusok/excel-exporter";
configureWasm({
  // Both optional; set only what you want to override.
  wasmUrl: "https://cdn.example.com/modern-xlsx.wasm",
  workerUrl: "https://cdn.example.com/export.worker.js",
});
```

Bundlers with asset imports can also wire the files explicitly (the pre-2.0 recommended setup — still fully supported):

```ts
import wasmUrl from "@marcusok/excel-exporter/dist/modern-xlsx.wasm?url";
import workerUrl from "@marcusok/excel-exporter/dist/export.worker.js?url";
configureWasm({ wasmUrl, workerUrl });
```

Without a bundler (plain `<script type="module">`), copy the two files out of this package's `dist/` into a static directory and point `configureWasm` at them.

### Multi-level Headers and Merges

Columns support a `children` tree to produce multi-row headers: a group column header automatically merges across all its leaf columns, and a leaf column header automatically spans the remaining header rows vertically — no manual merge-range math needed. Data-area merges use `merges` (positioned relative to the data area, where `row 0` is the first data row).

```ts
columns: [
  { prop: "product", label: "Product" },
  {
    label: "Revenue",
    children: [
      {
        label: "This month",
        children: [
          { prop: "m_qty", label: "Qty" },
          { prop: "m_amt", label: "Amount" },
        ],
      },
      {
        label: "Year to date",
        children: [
          { prop: "y_qty", label: "Qty" },
          { prop: "y_amt", label: "Amount" },
        ],
      },
    ],
  },
],
```

Multi-level headers and merges (including header merges) work on all routes — main / worker / stream (including the style-less stream fallback). The stream path preserves merges but does not support styles.

Invalid input fails identically on every route with `{ success: false, error }` instead of a corrupt workbook: `merges` must be integers (`row`/`col` ≥ 0, `rowspan`/`colspan` ≥ 1) staying within the data area and not overlapping each other; sheet names must be unique across `sheets`; `NaN`/`Infinity` in unformatted numeric columns are written as visible strings (illegal XML numbers would corrupt the file). Values are normalized identically on every route: without a `format`, plain objects are written as JSON strings, `Date`s as ISO strings and bigints as decimal strings, so a dataset crossing the 50k-row threshold keeps the same cell content.

### Auto Routing

`pickMode()` in `index.ts` selects the optimal path based on row count (overridable via the `mode` option):

| Rows               | Browser              | Node/SSR |
| ------------------ | -------------------- | -------- |
| < 20,000 rows      | main                 | main     |
| 20,000–49,999 rows | Worker + Workbook    | main     |
| >= 50,000 rows     | Worker + Fast stream | stream   |

On the Worker path the main thread only performs one structured-clone `postMessage` (100k rows ~94ms); the export work runs in the Worker thread. The Workbook path supports full `CellStyle`; Fast stream supports multi-row headers and merges but not `StyleBuilder`/layout styles — `width`/`freezeRows` etc. are dropped with a warning.

### Style Presets

[`src/style-presets.ts`](./src/style-presets.ts) provides 8 presets: `header` (bold, dark-blue background with white text), `currency` (thousands separator, two decimals), `date`/`datetime`, `percent`, `dataRow` (left-aligned, light-gray bottom border), `bordered` (thin light-gray box on all four sides), `danger` (red bold). Custom `CellStyle` is supported (font/fill/alignment/borders/number format); colors are 6-digit RGB hex (e.g. `'FF0000'`).

Presets are plain constant objects — combine them at the sheet level, and derive variants with object spread instead of rewriting a `CellStyle` from scratch:

```ts
// sheet-level fields (see "Table-wide styling & index column" below)
headerStyle: StylePresets.header, // headers, whole table
dataStyle: StylePresets.bordered, // base style for every data cell
// column-level: a tweaked preset — thousands separator, no decimals
style: { ...StylePresets.currency, numFormat: "#,##0" },
```

Variant recipes and the full styling guide (header overrides, index-column styling, format-code tweaks): see the [docs](https://yourbusiness.github.io/marcusok/packages/excel-exporter/guide/05-styles).

### Table-wide styling & index column

Two sheet-level fields cover the whole table in one place:

- **`dataStyle`** — base `CellStyle` for every data cell. A column's own `style` deep-merges over it field by field (a table-wide border survives a column that only sets `numFormat`, and vice versa), so `dataStyle: StylePresets.bordered` plus a few column tweaks is the idiomatic bordered-table setup. Headers stay with `headerStyle`.
- **`indexColumn`** — injects a leading row-number column (`true`, or `{ label, width, start, style, headerStyle }`): values come from the row number (never read from `data`), existing `merges` shift right automatically, and the feature works on every export path — workbook, worker and the style-less stream alike. The header label defaults to `序号` and the width to `6`. The expansion happens once inside `exportExcel`; the lower-level `WorkbookBuilder` / `exportAsStream` entry points only understand an already-expanded `__index__` column, so call the exported `applyIndexColumn(sheet)` before driving them directly.
- **Default alignment** — every cell is centered by default: the package's `BaseCellStyle` (horizontal + vertical `center`) sits underneath all cells, headers included. Explicit `alignment` values win field by field, and either `dataStyle` or `headerStyle` overrides the base; declare both axes to restore Excel's native alignment. Presets keep their own alignment — `StylePresets.dataRow` stays left-aligned by design.

```ts
sheets: [
  {
    name: "Sheet1",
    headerStyle: StylePresets.header, // headers
    dataStyle: StylePresets.bordered, // all data cells
    indexColumn: true, // header label defaults to "序号"
    columns: [
      { prop: "name", label: "Name" },
      { prop: "amount", label: "Amount", style: StylePresets.currency },
    ],
    data,
  },
];
```

### Value Formatting

See [`src/types.ts`](./src/types.ts). Worker mode cannot carry functions across structured clone, so a declarative `FormatSpec` is provided:

| Type       | Example                                                        | Description                     |
| ---------- | -------------------------------------------------------------- | ------------------------------- |
| `enum`     | `{ type: "enum", map: { paid: "Paid" }, fallback: "Unknown" }` | Enum mapping                    |
| `number`   | `{ type: "number", decimals: 2, thousands: true }`             | Number precision and separators |
| `date`     | `{ type: "date", pattern: "yyyy/MM/dd" }`                      | Date serialization              |
| `datetime` | `{ type: "datetime", pattern: "yyyy-MM-dd HH:mm:ss" }`         | Date-time serialization         |
| `padding`  | `{ type: "padding", fill: "0", length: 6, align: "left" }`     | String padding                  |

Main-thread paths additionally accept function form (`main`, and `stream` in Node where it runs on the main thread): `format: (v) => v ? "Yes" : "No"`. The browser Worker path strips functions and prints a warning — use `FormatSpec` there instead.

Number-spec cross-path notes (see `ColumnConfig.format` in [`src/types.ts`](./src/types.ts)):

- Set `decimals` explicitly: the Workbook path stores full precision and renders decimals via `numFormat`, while the stream path (>= 50,000 rows / degraded exports) bakes decimals into the stored value.
- `thousands: true` renders separators only on the Workbook path (`#,##0` numFormat). The stream path keeps the cell a number, so separators are not visible there — baking them into the value would turn data cells into text.
- `null`/`undefined` values render as empty cells on every path (never `0`).

### Fallback

When the browser Worker route fails (missing/404 worker asset, WASM init error inside the Worker, timeout), the recovery depends on which engine that route was running:

- **Worker + Workbook** (browser 20,000 – 49,999 rows): the library first **retries on the main thread** with modern-xlsx — styles are preserved. Only when that retry also fails (or WASM is unsupported / fails to load on the main thread) does the export degrade to the **pure-JS fast stream** — no WASM, no network, headers and merges preserved, styles stripped. A successful degraded export carries `result.error` (with `success: true`) describing the degradation.
- **Worker + Fast stream** (browser ≥ 50,000 rows, or an explicit `mode: "stream"`): the retry on the main thread _is_ the same fast stream on the identical input, so there are no styles left to preserve and nothing further to degrade to. A retry failure is terminal — the export resolves with `success: false` instead of attempting a doomed third build.

Either way the caller's promise resolves rather than rejects, and each degradation step prints an `[excel-exporter]` console warning — check `result.error` to monitor the fallback rate.

### Progress Overlay

An optional full-screen overlay with a progress bar, behind its own subpath so the main entry stays free of DOM code:

```ts
import { exportExcelWithOverlay } from "@marcusok/excel-exporter/overlay";

const result = await exportExcelWithOverlay(
  { filename: "report", sheets: [...] },
  { delayMs: 200, blockInteraction: true },
);
```

It appends to (never replaces) your existing `onProgress` / `onPhase` callbacks, blocks page interaction while shown, and is removed when the export settles — on success and on failure alike. Only the Fast stream path emits intermediate progress, so Workbook routes render an animated sweep instead of a percentage; `delayMs` keeps fast exports from flashing an overlay at all. See the [Progress Overlay guide](https://yourbusiness.github.io/marcusok/packages/excel-exporter/guide/11-overlay) for the route-by-route behaviour, the blocking-thread trade-offs, and the handle form used by `exportTable` / `exportEcharts`.

## API

- `exportExcel(options)` — unified entry with auto routing.
- `configureWasm(opts)` — optional overrides for `wasmUrl`/`workerUrl`/`timeoutMs`/`maxRetries`/`workerTimeoutMs` (see [How assets resolve](#how-assets-resolve-zero-configuration)). Note: changing `wasmUrl` after a _successful_ load does not reload WASM on a thread that already initialized it (modern-xlsx's `initWasm` is idempotent — first successful init wins); the new URL takes effect only in a fresh JS realm (page reload / a worker created after `terminateWorker()`), and a console warning is printed when this applies. Changing `workerUrl` has the analogous limitation: the shared Worker reads its script URL once at creation, so a later change only reaches a worker created after `terminateWorker()` (a warning is printed on any `workerUrl` change).
- `onPhase(phase, durationMs)` (an `exportExcel` option) — per-phase timing callback: `init` (WASM init) / `build` (workbook build) / `download` (trigger download); reports elapsed milliseconds per phase — note that each real build attempt reports its own `build` phase, so a degradation chain (failed worker build → main-thread retry → stream fallback) reports one `build` per attempt (see `ExportPhase` in [`src/types.ts`](./src/types.ts)). Does not affect the `duration` in the returned result.
- `WorkbookBuilder` — batch builder (<50k rows, full styling). Does not expand `SheetConfig.indexColumn`; call `applyIndexColumn(sheet)` first.
- `exportAsStream(sheets)` — large-file export (>=50k rows). Same caveat as `WorkbookBuilder` for `indexColumn`.
- `applyIndexColumn(sheet)` / `INDEX_PROP` — expand `indexColumn` into the leading `__index__` column yourself (what `exportExcel` does internally), and the reserved prop name, for callers driving the two low-level entry points above.
- `exportTable(options)` — convenience export for common table data, accepting Element Plus `prop`/`label` (the library naming), AntD `dataIndex`/`title`, and the legacy `key`/`header` names.
- `exportEcharts(options)` — convenience export for common ECharts data, supporting category-axis multi-series, pie `name/value`, and scatter pairs in either ECharts spelling (`[x,y]` or `{ value: [x,y] }`). The default sheet name and column headers are Chinese (`图表数据` / `系列` / `类目` / `名称` / `数值`), except the scatter layout, whose coordinate headers are the literal `X` / `Y`; override them via `sheetName` / `seriesHeader` / `categoryHeader` / `nameHeader` / `valueHeader`. In long/item layouts the header texts double as row keys, so duplicated headers are rejected with a clear error.
- `StylePresets` — the eight preset styles, also importable on their own from the `@marcusok/excel-exporter/styles` subpath.
- `headerStyle` — supported on both `SheetConfig` and `ColumnConfig` for styling header cells.
- `exportExcelWithOverlay(options, overlayOptions)` / `showExportOverlay(overlayOptions)` (`@marcusok/excel-exporter/overlay`, source entry `src/overlay.ts`) — optional full-screen progress overlay; the first wraps `exportExcel`, the second returns a handle for `exportTable` / `exportEcharts` / custom flows. No-ops in Node/SSR.
- `exportInWorker` / `terminateWorker` (`@marcusok/excel-exporter/worker-utils`, source entry `src/worker-exporter.ts`) — manual Worker lifecycle control.

## Node Usage

Node has no Web Worker, so auto routing degrades to main (<50k rows) or stream (>=50k rows) on the main thread.

**No boilerplate needed.** With nothing configured, the engine reads this package's `dist/modern-xlsx.wasm` from disk (relative to the installed package, pnpm-symlink-safe) and initializes it synchronously (`initWasmSync`) on first use — there is nothing to call and nothing to copy.

To control initialization timing yourself (e.g. move the one-off synchronous read+compile to startup instead of the first request), await the loader once at boot:

```ts
import { getWasmLoader } from "@marcusok/excel-exporter";
await getWasmLoader().ensureLoaded(); // reads + compiles the shipped wasm once
```

Alternatively, `configureWasm({ wasmUrl })` with an HTTP(S) URL works too (fetched, not read from disk). Note: `initWasmSync` from a separately installed `modern-xlsx` does **not** pre-warm this package — the engine is bundled in, so an external copy is a different module instance.

## Design Decisions

- **50k-row cutover**: `STREAM_THRESHOLD=50_000` (branch `>=`); below 50k rows uses Workbook (full styling), 50k and above uses Fast stream.
- **Worker threshold at 20,000 rows**: below 20k rows uses main (10k×6 columns measures ~120ms in a browser); 20k and above uses a Worker to avoid long main-thread blocking.
- **Zero runtime dependencies**: the engine (modern-xlsx glue, fflate) is bundled at build time and the wasm binary ships under this package's own `exports` map — consumers install one package, and upstream `engines` declarations never leak into their install.
- **Self-contained worker**: `dist/export.worker.js` is a single ESM file with zero imports. Bundlers emit it verbatim as an asset (the default `new URL` resolution), so a chunked worker whose sibling imports are not tracked can never 404 in production builds.
- **ESM-only**: this package provides no CJS build.
- **Worker-compatible format**: functions cannot cross structured clone. The browser Worker path (including stream executed inside a Worker) accepts only `FormatSpec`, and `exportInWorker` strips function formats; Node's stream runs on the main thread, so functions are fine there.
- **Fast stream has no styles**: the large-file path emits minimal OOXML and does not support `StyleBuilder`. Multi-row headers and merges are preserved; `width`/`freezeRows` etc. are dropped with a warning under stream.
- **Concurrency safety**: Worker communication routes by requestId with a `pending: Map`; `onmessage` is registered only once.

## License

MIT
