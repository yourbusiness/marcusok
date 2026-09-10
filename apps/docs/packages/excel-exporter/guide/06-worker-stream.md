# Worker & Streaming

## Worker threading (≥ 20,000 rows)

At ≥ 20,000 rows in the browser, `auto` selects the Worker path: the main thread only does one structured clone (~94ms at 100k rows) and everything else happens inside the Worker. Rows 20,000–49,999 load WASM and build inside the Worker (Workbook path); at ≥ 50,000 rows it switches to the WASM-free Fast stream (see below).

The worker asset (`export.worker.js`, a self-contained ESM file) is located automatically — nothing to configure. Set `configureWasm({ workerUrl })` only to point at a self-hosted copy.

Worker path behavior:

- **Worker failures degrade gracefully** — if the Worker route fails (e.g. a misconfigured `workerUrl` override 404s), the export **retries on the main thread** (modern-xlsx keeps styles; the ≥ 50,000-row fast stream needs no WASM at all). Only if that retry also fails does it degrade to the style-less stream fallback (`mode: "stream"` with `result.error` set, success stays `true`). The caller's promise resolves rather than rejects, with an `[excel-exporter]` console warning at each degradation step;
- The Worker instance is reused and requests are dispatched by `requestId`, so concurrent exports never interfere;
- **Function-form formats are stripped** (functions cannot be structured-cloned) — use FormatSpec on worker paths;
- `onProgress` / `onPhase` (`init` / `build`) are forwarded from the Worker.

## Streaming writes (≥ 50,000 rows)

`fast-xlsx.ts` uses `fflate` to produce minimal OOXML: ~0.8s at 100k rows (vs 17.5s on the Workbook path). `auto` selects it at ≥ 50k rows.

Known stream limitations (v1):

| Feature                        | Stream path                                        |
| ------------------------------ | -------------------------------------------------- |
| Multi-row headers (`children`) | supported (header auto-merge)                      |
| Cell merges (`merges`)         | supported (data area)                              |
| Cell styles (`style`)          | not supported                                      |
| Header styles (`headerStyle`)  | not supported                                      |
| Column width (`width`)         | not supported                                      |
| Freeze / auto-filter           | not supported                                      |
| Custom number formats          | not supported (`decimals` baked into stored value) |
| Date formats                   | readable strings per pattern                       |
| Progress callback              | reported every 1000 rows                           |

Skipped features (cell styles, header styles, column width, freeze, ...) print `[excel-exporter] stream mode: features not supported (...)` in the console.

## Which path should I use?

| Need                            | Recommended                       |
| ------------------------------- | --------------------------------- |
| < 50k rows with full styling    | `auto` (main / worker + Workbook) |
| ≥ 50k rows, styling can degrade | `auto` (worker + Fast stream)     |
| Large batch in Node             | `auto` (main → stream at ≥ 50k)   |
| Zero main-thread blocking       | explicit `mode: "worker"`         |

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
