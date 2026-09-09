# @marcusok/excel-exporter · Excel Export Engine

An Excel export library built on [modern-xlsx](https://github.com/ABCrimson/modern-xlsx) (WASM) plus a custom Fast stream writer. It offers a declarative API, automatic mode routing, full cell styling, Web Worker multithreading, fast writes, and a SheetJS fallback.

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

Environment: Node >= 22 (any package manager works — the examples here use pnpm; `pnpm >= 9` is only a requirement of this repo's own development setup). The engine dependency `modern-xlsx@^1.2.0` is installed automatically. modern-xlsx@1.2.0 declares `engines.node>=24`, but its WASM core targets browsers; this package's full test suite passes on Node 22 (CI runs there). If your package manager enforces engines checks, installation fails on Node 22 — set `engine-strict=false` in your project's `.npmrc` (the same approach this repository uses) or upgrade to Node >= 24.

> modern-xlsx is a direct `dependency`: this package pins the engine version it was tested against, and re-publishes `modern-xlsx.wasm` under its own `exports` map (`@marcusok/excel-exporter/dist/modern-xlsx.wasm`), so the binary always ships with the matching JS glue — and bundlers can import it directly (modern-xlsx's own `exports` map omits its wasm subpaths, which would otherwise force a manual copy step). `xlsx` (SheetJS) is an optional peerDep, needed only for the fallback path.
>
> Security note on the optional `xlsx` peer: the last npm release (`0.18.5`) is unmaintained and carries known CVEs (CVE-2023-30533 ReDoS, CVE-2024-22363 prototype pollution). If you provide `xlsx` yourself, install the maintained build from the official CDN instead of npm: `pnpm add https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`. The peer range stays `>=0.18.5` for compatibility, and when no local `xlsx` is present the fallback loads `0.20.3` from the SheetJS CDN at runtime.

## Setup (Browser)

One asset is required on nearly every route: `modern-xlsx.wasm` (1.9MB, the style engine — the only exception is an explicit `mode: "stream"`, whose pure-JS fast stream needs no WASM). A second one, `export.worker.js`, is needed only when exports actually enter a Worker — auto mode with >= 20,000 rows, or an explicit `mode: "worker"` / `mode: "stream"`. Both files live in this package's `dist/` (the wasm is forwarded there at build time), so everything resolves through `@marcusok/excel-exporter` imports.

### Vite (recommended)

Import the assets with the `?url` suffix — Vite serves them in dev and hashes them into `dist/assets/` at build time. No copy plugin, no `public/` directory, nothing hardcoded:

```ts
// main.ts
import { configureWasm } from "@marcusok/excel-exporter";
// Re-published by this package (modern-xlsx's own exports map omits wasm subpaths)
import wasmUrl from "@marcusok/excel-exporter/dist/modern-xlsx.wasm?url";
// Optional — drop this line entirely if your exports stay under 20k rows
import workerUrl from "@marcusok/excel-exporter/dist/export.worker.js?url";

configureWasm({ wasmUrl, workerUrl });
```

### Other setups (copy the files)

If your bundler has no asset-URL import, copy both files out of this package's `dist/` into a static directory. Resolving the real path via `require.resolve` in `buildStart` avoids hardcoded `node_modules` paths (which break under pnpm symlinks):

```ts
// vite.config.ts / build script
import { defineConfig } from "vite";
import { createRequire } from "node:module";
import { copyFileSync, mkdirSync, statSync } from "node:fs";
import { dirname } from "node:path";

const require = createRequire(import.meta.url);
// Single source: both assets ship in @marcusok/excel-exporter/dist
const pkgDist = dirname(require.resolve("@marcusok/excel-exporter"));

export default defineConfig({
  plugins: [
    {
      name: "copy-excel-exporter-assets",
      buildStart() {
        mkdirSync("public/assets", { recursive: true });
        for (const file of ["modern-xlsx.wasm", "export.worker.js"]) {
          const src = `${pkgDist}/${file}`;
          if (!statSync(src, { throwIfNoEntry: false }))
            throw new Error(`${file} not found. Looked at: ${src}`);
          copyFileSync(src, `public/assets/${file}`);
        }
      },
    },
  ],
});
```

```ts
// main.ts
import { configureWasm } from "@marcusok/excel-exporter";
configureWasm({
  wasmUrl: "/assets/modern-xlsx.wasm",
  workerUrl: "/assets/export.worker.js",
});
```

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

Multi-level headers and merges (including header merges) work on all four paths — main / worker / stream / SheetJS fallback. The stream and fallback paths preserve merges but still do not support styles.

Invalid input fails identically on every path with `{ success: false, error }` instead of a corrupt workbook: `merges` must be integers (`row`/`col` ≥ 0, `rowspan`/`colspan` ≥ 1) staying within the data area and not overlapping each other; sheet names must be unique across `sheets`; `NaN`/`Infinity` in unformatted numeric columns are written as visible strings (illegal XML numbers would corrupt the file). Values are normalized identically on every path: without a `format`, plain objects are written as JSON strings, `Date`s as ISO strings and bigints as decimal strings, so a dataset crossing the 50k-row threshold keeps the same cell content.

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

- Set `decimals` explicitly: the Workbook path stores full precision and renders decimals via `numFormat`, while the stream/SheetJS paths (>= 50,000 rows / degraded exports) bake decimals into the stored value.
- `thousands: true` renders separators only on the Workbook path (`#,##0` numFormat). The stream/SheetJS paths keep the cell a number, so separators are not visible there — baking them into the value would turn data cells into text.
- `null`/`undefined` values render as empty cells on every path (never `0`).

### Fallback

When the browser Worker route fails (missing/404 `workerUrl`, WASM init error inside the Worker, timeout), the library first **retries on the main thread** with modern-xlsx — styles are preserved, and the ≥ 50,000-row fast stream needs no WASM at all. Only when that retry also fails (or WASM is unsupported / fails to load on the main thread) does the export degrade to SheetJS ([`src/fallback.ts`](./src/fallback.ts)); fallback exports carry no styles. `ExportResult.engine` reports `'sheetjs'` so you can monitor the fallback rate, and each degradation step prints an `[excel-exporter]` console warning.

## API

- `exportExcel(options)` — unified entry with auto routing.
- `configureWasm(opts)` — set `wasmUrl`/`workerUrl`/`timeoutMs`/`maxRetries`. Note: changing `wasmUrl` after a _successful_ load does not reload WASM on a thread that already initialized it (modern-xlsx's `initWasm` is idempotent — first successful init wins); the new URL takes effect only in a fresh JS realm (page reload / a worker created after `terminateWorker()`), and a console warning is printed when this applies.
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

**No boilerplate needed.** When no `wasmUrl` is configured, the engine locates `modern-xlsx.wasm` through `node_modules` on first use (`createRequire`, pnpm-symlink-safe) and initializes it synchronously (`initWasmSync`). Node's `fetch` rejects the `file://` auto-detected URL, so this path replaces the manual init snippet previous versions required — there is nothing to call and nothing to copy.

To control initialization timing yourself (e.g. move the one-off synchronous read+compile to startup instead of the first request), keep the explicit form, which remains fully supported:

```ts
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { initWasmSync } from "modern-xlsx";
const require = createRequire(import.meta.url);
initWasmSync(
  readFileSync(
    `${require("path").dirname(require.resolve("modern-xlsx"))}/modern-xlsx.wasm`,
  ),
);
```

Alternatively, `configureWasm({ wasmUrl })` with an HTTP(S) URL works too (fetched, not read from disk). Node version: this package declares `engines.node >=22`, and CI runs Node 22. The dependency modern-xlsx declares `>=24`, but its WASM core targets browsers — everything is green on Node 22.

## Design Decisions

- **50k-row cutover**: `STREAM_THRESHOLD=50_000` (branch `>=`); below 50k rows uses Workbook (full styling), 50k and above uses Fast stream.
- **Worker threshold at 20,000 rows**: below 20k rows uses main (10k×6 columns measures ~120ms in a browser); 20k and above uses a Worker to avoid long main-thread blocking.
- **ESM-only**: modern-xlsx ships ESM only, and this package provides no CJS build.
- **Worker-compatible format**: functions cannot cross structured clone. The browser Worker path (including stream executed inside a Worker) accepts only `FormatSpec`, and `exportInWorker` strips function formats; Node's stream runs on the main thread, so functions are fine there.
- **Fast stream has no styles**: the large-file path emits minimal OOXML and does not support `StyleBuilder`. Multi-row headers and merges are preserved; `width`/`freezeRows` etc. are dropped with a warning under stream.
- **Concurrency safety**: Worker communication routes by requestId with a `pending: Map`; `onmessage` is registered only once.

## License

MIT
