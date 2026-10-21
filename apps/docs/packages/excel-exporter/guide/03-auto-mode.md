# Auto Mode Routing

The `mode` option defaults to `"auto"`: the library picks the path from **total row count and estimated cell count** plus the runtime environment, so business code never needs to react to data size. Cell count = data rows × effective leaf columns (after `indexColumn` expansion), summed across all sheets of one call.

## Routing rules

### Browser

| Workload                                                                                 | Path                   | Notes                                                                                  |
| ---------------------------------------------------------------------------------------- | ---------------------- | -------------------------------------------------------------------------------------- |
| `< 20,000` rows **and** `< 250,000` cells                                                | `main`                 | Synchronous main-thread build; ~120ms at 10k×6 columns                                 |
| rows `>= 20,000` **or** cells `>= 250,000`, with rows `< 50,000` and cells `< 1,500,000` | `worker` + Workbook    | Main thread does one structured clone (grows with cell count); WASM runs in the Worker |
| rows `>= 50,000` **or** cells `>= 1,500,000`                                             | `worker` + Fast stream | Custom fflate writes avoid the `toBuffer` cliff                                        |

### Node / SSR (no Web Worker)

| Workload                                     | Path     | Notes                                    |
| -------------------------------------------- | -------- | ---------------------------------------- |
| rows `< 50,000` **and** cells `< 1,500,000`  | `main`   | Main-thread Workbook build, full styling |
| rows `>= 50,000` **or** cells `>= 1,500,000` | `stream` | Fast stream, ~0.8s at 100k rows          |

## Explicit modes

```ts
await exportExcel({ ..., mode: "stream" });  // force streaming
await exportExcel({ ..., mode: "worker" });  // force Worker (browser)
await exportExcel({ ..., mode: "main" });    // force main thread
```

`mode: "worker"` does **not** error in Node/SSR: it falls back to the main-thread path (stream above the stream gate), preserving style semantics instead of silently degrading to the style-less stream. In a browser, an explicit `worker` route uses the same row/cell gate to decide between Workbook and Fast stream inside the Worker.

## Why 20,000 / 50,000 / cell budgets

- **20,000 rows or 250,000 cells**: synchronous main-thread work is acceptable below this level; above it the browser hands the build to a Worker to keep the UI responsive.
- **50,000 rows or 1,500,000 cells**: `Workbook.toBuffer()` shows a superlinear cliff beyond ~55k rows (~17.5s at 100k in Chrome; the same cliff measures ~21s cold in Node — see [Performance](/packages/excel-exporter/guide/07-performance)), while Fast stream stays at ~0.8s. The cell gate closes the other blind spot: a wide sheet can be far below 50k rows and still be expensive on the Workbook path. Node measurements: 40k×50 (2M cells) ~6.3s / RSS +1.2GB on Workbook vs ~1s / +58MB on stream; 20k×100 (2M cells) ~3.2s / +258MB vs stream; 60k×20 (1.2M cells) ~1s on stream.

> Trade-off: the stream path supports multi-row headers, cell merges and number formats (date/number FormatSpec via minimal `styles.xml`), but not cell styles, header styles or layout features (width/freeze/filter). Auto mode routes on the **total** rows and cells across all sheets of one call, so splitting the data into more sheets of the same call does not preserve styling. For fully-styled exports keep both the per-call total under the gates, or split into multiple `exportExcel` calls (one file each).
>
> When the stream path drops features, `ExportResult.degradedFeatures` lists them (`width`, `headerStyle`, `style`, `dataStyle`, `freezeRows`, `autoFilter`). The field is set on the normal stream route even though `success` is `true` and `error` is unset; a failure-driven fallback also sets `error` with the reason.
