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
pnpm add @marcusok/excel-exporter modern-xlsx
```

Environment: Node >= 22 (any package manager works — the examples here use pnpm; `pnpm >= 9` is only a requirement of this repo's own development setup). modern-xlsx@1.2.0 declares `engines.node>=24`, but its WASM core targets browsers; this package passes all tests on Node 22 (84 cases in total; CI defaults to `RUN_PERF=0`, skipping 4 performance benchmarks and running 80). This package was developed and tested against 1.2.0 — consumers are advised to pin that version (the peerDep range `^1.2.0` is allowed, but higher versions are unverified).

> modern-xlsx is declared as a `peerDependency`, so consumers must install it explicitly. Reasons: (1) `modern-xlsx.wasm` (1.9MB) must be deployed by the consumer as a static asset — an implicit dependency would hide this hard requirement; (2) peerDep is semantically correct — this package wraps modern-xlsx and version control belongs to the consumer; (3) package managers auto-install peerDependencies by default (npm 7+ / pnpm 8+), and an implicitly installed version is outside the consumer's control — an explicit declaration is what pins the version intent. `xlsx` (SheetJS) is an optional peerDep, needed only for the fallback path.

## Setup (Browser)

Two static assets must be reachable from the consuming site: `modern-xlsx.wasm` (1.9MB) and `export.worker.js`.

The recommended approach is a Vite plugin that resolves the real paths via `require.resolve` in `buildStart` and copies them into `public/assets/`, avoiding hardcoded `node_modules` paths (incompatible with pnpm symlinks). See [design doc 6.2](https://github.com/yourbusiness/marcusok/blob/main/docs/excel-export-design.md) for details.

```ts
// vite.config.ts
import { defineConfig } from "vite";
import { createRequire } from "node:module";
import { copyFileSync, mkdirSync, statSync } from "node:fs";
import { dirname } from "node:path";

const require = createRequire(import.meta.url);
const resolveDistDir = (spec: string) => dirname(require.resolve(spec));

export default defineConfig({
  plugins: [
    {
      name: "copy-modern-xlsx-assets",
      buildStart() {
        mkdirSync("public/assets", { recursive: true });
        copyFileSync(
          `${resolveDistDir("modern-xlsx")}/modern-xlsx.wasm`,
          "public/assets/modern-xlsx.wasm",
        );
        const workerSrc = `${resolveDistDir("@marcusok/excel-exporter")}/export.worker.js`;
        if (!statSync(workerSrc, { throwIfNoEntry: false }))
          throw new Error(
            `export.worker.js not found. Run pnpm build first. Looked at: ${workerSrc}`,
          );
        copyFileSync(workerSrc, "public/assets/export.worker.js");
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

Invalid input fails identically on every path with `{ success: false, error }` instead of a corrupt workbook: `merges` must be integers (`row`/`col` ≥ 0, `rowspan`/`colspan` ≥ 1) staying within the data area and not overlapping each other; sheet names must be unique across `sheets`; `NaN`/`Infinity` in unformatted numeric columns are written as visible strings (illegal XML numbers would corrupt the file).

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

### Fallback

When WASM is unsupported or fails to load, the library automatically falls back to SheetJS ([`src/fallback.ts`](./src/fallback.ts)); fallback exports carry no styles. `ExportResult.engine` reports `'sheetjs'` so you can monitor the fallback rate.

## API

- `exportExcel(options)` — unified entry with auto routing.
- `configureWasm(opts)` — set `wasmUrl`/`workerUrl`/`timeoutMs`/`maxRetries`.
- `onPhase(phase, durationMs)` (an `exportExcel` option) — per-phase timing callback: `init` (WASM init) / `build` (workbook build) / `download` (trigger download); reports elapsed milliseconds once per phase for metrics breakdowns, without affecting the `duration` in the returned result.
- `WorkbookBuilder` — batch builder (<50k rows, full styling).
- `exportAsStream(sheets)` — large-file export (>=50k rows).
- `exportTable(options)` — convenience export for common table data, supporting both AntD `title`/`dataIndex` and Element Plus `label`/`prop` column naming.
- `exportEcharts(options)` — convenience export for common ECharts data, supporting category-axis multi-series, pie `name/value`, and scatter `[x,y]`.
- `StylePresets` — the seven preset styles.
- `headerStyle` — supported on both `SheetConfig` and `ColumnConfig` for styling header cells.
- `exportInWorker` / `terminateWorker` (`@marcusok/excel-exporter/worker-utils`, source entry `src/worker-exporter.ts`) — manual Worker lifecycle control.

## Node Usage

Node has no Web Worker, so auto routing degrades to main (<50k rows) or stream (>=50k rows) on the main thread.

Node's `fetch` rejects the `file://` protocol, so **when consuming the package locally in Node you cannot rely on `exportExcel()` to auto-load WASM**. A production server must initialize WASM explicitly first (`initWasmSync`), or provide a fetchable HTTP URL via `configureWasm({ wasmUrl })`:

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

Node version: this package declares `engines.node >=22`, and CI runs Node 22. The peer modern-xlsx declares `>=24`, but its WASM core targets browsers — everything is green on Node 22.

## Design Decisions

- **50k-row cutover**: `STREAM_THRESHOLD=50_000` (branch `>=`); below 50k rows uses Workbook (full styling), 50k and above uses Fast stream.
- **Worker threshold at 20,000 rows**: below 20k rows uses main (10k×6 columns measures ~120ms in a browser); 20k and above uses a Worker to avoid long main-thread blocking.
- **ESM-only**: modern-xlsx ships ESM only, and this package provides no CJS build.
- **Worker-compatible format**: functions cannot cross structured clone. The browser Worker path (including stream executed inside a Worker) accepts only `FormatSpec`, and `exportInWorker` strips function formats; Node's stream runs on the main thread, so functions are fine there.
- **Fast stream has no styles**: the large-file path emits minimal OOXML and does not support `StyleBuilder`. Multi-row headers and merges are preserved; `width`/`freezeRows` etc. are dropped with a warning under stream.
- **Concurrency safety**: Worker communication routes by requestId with a `pending: Map`; `onmessage` is registered only once.

## License

MIT
