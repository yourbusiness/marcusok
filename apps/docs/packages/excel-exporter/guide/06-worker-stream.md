# Worker & Streaming

## Worker threading (≥ 20,000 rows or ≥ 250,000 cells)

At ≥ 20,000 rows **or ≥ 250,000 estimated cells** in the browser, `auto` selects the Worker path: the main thread only does one structured clone and everything else happens inside the Worker. The clone cost grows with cell count (~94ms at 100k×6, ~200ms at 40k×50 in Chromium). Below the stream gate (rows < 50,000 **and** cells < 1,500,000), the Worker loads WASM and builds with the Workbook path; at rows ≥ 50,000 **or** cells ≥ 1,500,000 it switches to the WASM-free Fast stream (see below).

The worker asset (`export.worker.js`, a self-contained ESM file) is located automatically — nothing to configure. Set `configureWasm({ workerUrl })` only to point at a self-hosted copy.

Worker path behavior:

- **Worker failures degrade gracefully** — if the Worker route fails (e.g. a misconfigured `workerUrl` override 404s), the export **retries on the main thread** (modern-xlsx keeps styles; the stream-gate fast stream needs no WASM at all). On the Worker + Workbook route a failing retry degrades further to the style-less stream fallback (`mode: "stream"` with `result.error` set, success stays `true`). On the Worker + stream route (browser at/above the stream gate) the retry _is_ the fast stream on the identical input, so a failure there is terminal: the export resolves with `success: false` instead of attempting a third build. Either way the caller's promise resolves rather than rejects, with an `[excel-exporter]` console warning at each degradation step;
- The Worker instance is reused, but requests are **serialized through a main-thread FIFO queue**: only one export is in flight at a time. Each request's timeout starts when it is dispatched (not while queued), and a timeout/abort only settles that request — queued exports continue on a fresh worker;
- **Function-form formats are stripped** (functions cannot be structured-cloned) — use FormatSpec on worker paths;
- `onProgress` / `onPhase` are forwarded from the Worker: `build` once the worker's build finishes, and `init` only on Worker + Workbook, when the worker actually initializes WASM — the Worker + stream route uses no WASM and reports no `init` at all.

## Streaming writes (≥ 50,000 rows or ≥ 1,500,000 cells)

`fast-xlsx.ts` uses `fflate` to produce minimal OOXML: ~0.8s at 100k rows (vs 17.5s on the Workbook path in Chrome — ~21s cold in Node; see [Performance](/packages/excel-exporter/guide/07-performance)). `auto` selects it at rows ≥ 50k **or** cells ≥ 1.5M.

Known stream limitations:

| Feature                        | Stream path                                                                           |
| ------------------------------ | ------------------------------------------------------------------------------------- |
| Multi-row headers (`children`) | supported (header auto-merge)                                                         |
| Cell merges (`merges`)         | supported (data area)                                                                 |
| Cell styles (`style`)          | not supported                                                                         |
| Header styles (`headerStyle`)  | not supported                                                                         |
| Column width (`width`)         | not supported                                                                         |
| Freeze / auto-filter           | not supported                                                                         |
| FormatSpec number/date formats | supported via minimal `styles.xml` (numFmt only; `CellStyle.numFormat` is dropped)    |
| Date formats                   | stored as serials + numFmt (same code as Workbook unless `style.numFormat` conflicts) |
| Progress callback              | reported every 1000 rows                                                              |

Skipped features (cell styles, header styles, column width, freeze, ...) print `[excel-exporter] stream mode: features not supported (...)` in the console and are returned as `result.degradedFeatures` — including on the normal auto route, where `success` stays `true` and `error` is unset. A failure-driven fallback also sets `error` with the degradation reason.

## Which path should I use?

| Need                                                             | Recommended                        |
| ---------------------------------------------------------------- | ---------------------------------- |
| Below both gates (rows < 50k and cells < 1.5M) with full styling | `auto` (main / worker + Workbook)  |
| Rows ≥ 50k or cells ≥ 1.5M, styling can degrade                  | `auto` (worker + Fast stream)      |
| Large batch in Node                                              | `auto` (main → stream at the gate) |
| Zero main-thread blocking                                        | explicit `mode: "worker"`          |

## Lower-level APIs

`WorkbookBuilder` and `exportAsStream` are exported for fine-grained control:

```ts
import { WorkbookBuilder, exportAsStream } from "@marcusok/excel-exporter";

// batch build
const builder = await WorkbookBuilder.create();
builder.addSheet(sheetA).addSheet(sheetB);
const bytes = await builder.toBuffer();

// streaming
const { bytes, rowCount } = await exportAsStream(sheets, onProgress);
```

::: warning
These lower-level entries only recognize an **already-expanded** `__index__`
column — unlike `exportExcel`, they do not expand `SheetConfig.indexColumn`
themselves, so an `indexColumn` on a sheet passed to them directly is silently
ignored. Call `applyIndexColumn(sheet)` first if you need the row-number column
on these paths (the same helper `exportExcel` uses internally). See
[`applyIndexColumn` in the API reference](/packages/excel-exporter/api/01-export-excel)
for details.
:::
