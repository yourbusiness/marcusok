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

That is the entire setup. The package has **zero runtime dependencies** — the engine (modern-xlsx JS glue + fflate) is bundled in at build time, and the 1.9MB `modern-xlsx.wasm` binary ships under this package's own `exports` map (`@marcusok/excel-exporter/dist/modern-xlsx.wasm`), so the binary always matches the JS glue. No engine package to install, no `engine-strict` conflicts from upstream `engines` ranges, no fallback package, nothing to configure in `main.ts` or your bundler.

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
        { key: "orderId", header: "Order ID", width: 18 },
        {
          key: "amount",
          header: "Amount",
          width: 12,
          style: StylePresets.currency,
        },
        {
          key: "status",
          header: "Status",
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

## How assets resolve (zero configuration)

Two files ship alongside the code and are located automatically:

- **`modern-xlsx.wasm`** (1.9MB, the style engine) — needed on every route except an explicit `mode: "stream"`, whose pure-JS fast stream needs no WASM at all.
- **`export.worker.js`** (self-contained, zero imports of its own) — needed only when exports actually enter a Worker (auto mode with ≥ 20,000 rows, or an explicit `mode: "worker"` / `mode: "stream"`).

Both default to `new URL(<file>, import.meta.url)` relative to the package entry:

- **Bundlers** (Vite dev's dependency pre-bundling and production builds — verified on Vite 8; webpack 5 documents the same `new URL(..., import.meta.url)` asset pattern) rewrite the expression and emit the file as a hashed asset. Nothing to import, copy or configure.
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
  { key: "product", header: "Product" },
  {
    header: "Revenue",
    children: [
      {
        header: "This month",
        children: [
          { key: "m_qty", header: "Qty" },
          { key: "m_amt", header: "Amount" },
        ],
      },
      {
        header: "Year to date",
        children: [
          { key: "y_qty", header: "Qty" },
          { key: "y_amt", header: "Amount" },
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

[`src/style-presets.ts`](./src/style-presets.ts) provides 7 presets: `header` (bold, dark-blue background with white text), `currency` (thousands separator, two decimals), `date`/`datetime`, `percent`, `dataRow` (left-aligned, light-gray bottom border), `danger` (red bold). Custom `CellStyle` is supported (font/fill/alignment/borders/number format); colors are 6-digit RGB hex (e.g. `'FF0000'`).

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

When the browser Worker route fails (missing/404 worker asset, WASM init error inside the Worker, timeout), the library first **retries on the main thread** with modern-xlsx — styles are preserved. Only when that retry also fails (or WASM is unsupported / fails to load on the main thread) does the export degrade to the **pure-JS fast stream** — no WASM, no network, headers and merges preserved, styles stripped. A successful degraded export carries `result.error` (with `success: true`) describing the degradation, and each degradation step prints an `[excel-exporter]` console warning — check `result.error` to monitor the fallback rate.

## API

- `exportExcel(options)` — unified entry with auto routing.
- `configureWasm(opts)` — optional overrides for `wasmUrl`/`workerUrl`/`timeoutMs`/`maxRetries` (see [How assets resolve](#how-assets-resolve-zero-configuration)). Note: changing `wasmUrl` after a _successful_ load does not reload WASM on a thread that already initialized it (modern-xlsx's `initWasm` is idempotent — first successful init wins); the new URL takes effect only in a fresh JS realm (page reload / a worker created after `terminateWorker()`), and a console warning is printed when this applies.
- `onPhase(phase, durationMs)` (an `exportExcel` option) — per-phase timing callback: `init` (WASM init) / `build` (workbook build) / `download` (trigger download); reports elapsed milliseconds once per phase for metrics breakdowns, without affecting the `duration` in the returned result.
- `WorkbookBuilder` — batch builder (<50k rows, full styling).
- `exportAsStream(sheets)` — large-file export (>=50k rows).
- `exportTable(options)` — convenience export for common table data, supporting both AntD `title`/`dataIndex` and Element Plus `label`/`prop` column naming.
- `exportEcharts(options)` — convenience export for common ECharts data, supporting category-axis multi-series, pie `name/value`, and scatter `[x,y]`. The default sheet name and column headers are Chinese (`图表数据` / `系列` / `类目` / `名称` / `数值`); override them via `sheetName` / `seriesHeader` / `categoryHeader` / `nameHeader` / `valueHeader`. In long/item layouts the header texts double as row keys, so duplicated headers are rejected with a clear error.
- `StylePresets` — the seven preset styles.
- `headerStyle` — supported on both `SheetConfig` and `ColumnConfig` for styling header cells.
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
