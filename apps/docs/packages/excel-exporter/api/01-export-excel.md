# API: exportExcel & Configuration

## exportExcel

```ts
exportExcel(options: ExportOptions): Promise<ExportResult>
```

The core entry point (convenience wrappers such as `exportTable` / `exportEcharts` delegate to it). Routes to main / worker / stream by row count and environment, degrading to a style-less fast stream when WASM is unavailable.

## ExportOptions

| Field        | Type                                               | Required | Description                                                                                                                                                                                                                                       |
| ------------ | -------------------------------------------------- | -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `sheets`     | `SheetConfig[]`                                    | yes      | At least one sheet                                                                                                                                                                                                                                |
| `filename`   | `string`                                           | yes      | Download name; `.xlsx` is appended unless it already ends with it                                                                                                                                                                                 |
| `mode`       | `"auto" \| "main" \| "worker" \| "stream"`         | —        | Default `"auto"`                                                                                                                                                                                                                                  |
| `onProgress` | `(progress: number) => void`                       | —        | 0 → 1; the leading 0 and trailing 1 are each fired exactly once by `exportExcel` on every route (including the stream fallback and ultimately failed exports); incremental progress only on the stream path (every 1,000 rows)                    |
| `onPhase`    | `(phase: ExportPhase, durationMs: number) => void` | —        | `init` / `build` / `download` timings                                                                                                                                                                                                             |
| `download`   | `boolean`                                          | —        | Default `true`; `false` returns the Blob only                                                                                                                                                                                                     |
| `overlay`    | `boolean \| ProgressOverlayOptions`                | —        | Default **on** since 2.8.0. `false` disables the full-screen progress overlay entirely; an options object customizes its texts, delay and theme. Node/SSR is a no-op. See the [Progress Overlay guide](/packages/excel-exporter/guide/11-overlay) |

## ExportResult

| Field       | Type            | Description                                    |
| ----------- | --------------- | ---------------------------------------------- |
| `success`   | `boolean`       | Whether the export succeeded                   |
| `blob?`     | `Blob`          | The file content                               |
| `engine?`   | `"modern-xlsx"` | Engine actually used                           |
| `mode?`     | `ExportMode`    | Mode actually used                             |
| `duration?` | `number`        | Total duration in ms                           |
| `rowCount?` | `number`        | Exported row count                             |
| `error?`    | `Error`         | Failure reason (also set on the fallback path) |

## configureWasm

```ts
configureWasm(options: LoaderOptions): void
```

Optional — assets default to the files shipped next to the package entry (see [Installation](/packages/excel-exporter/guide/02-installation)). Override for self-hosted copies, a CDN, or bundlers without asset-URL support.

| Field             | Default                 | Description                                                      |
| ----------------- | ----------------------- | ---------------------------------------------------------------- |
| `wasmUrl`         | the shipped `.wasm`     | Override for a self-hosted / CDN copy                            |
| `workerUrl`       | the shipped worker file | Override for a self-hosted / CDN copy                            |
| `timeoutMs`       | `10_000`                | Per-attempt load timeout                                         |
| `maxRetries`      | `3`                     | Max load attempts (3 total incl. the first; 300ms/600ms backoff) |
| `workerTimeoutMs` | `120_000`               | Worker export timeout                                            |

## Other exported symbols

Everything below comes from the main entry. Grouped by what you reach for it for.

### Low-level build entries

- `WorkbookBuilder` — `static create(): Promise<WorkbookBuilder>` · `addSheet(config: SheetConfig): this` · `toBuffer(): Promise<Uint8Array>` · `toBlob(): Promise<Blob>`. Batch build with full styling, no mode routing. Does **not** expand `SheetConfig.indexColumn` — pass an already-expanded sheet (see `applyIndexColumn` below), otherwise the field is silently ignored.
- `exportAsStream(sheets, onProgress?)` — `(sheets: SheetConfig[], onProgress?: (p: number) => void) => Promise<{ bytes: Uint8Array; rowCount: number }>`. The style-less fast stream, called directly. The second argument is the callback itself, **not** an options object. Same `indexColumn` caveat as `WorkbookBuilder`.
- `applyIndexColumn(sheet)` / `INDEX_PROP` — `(sheet: SheetConfig) => SheetConfig`, `INDEX_PROP = "__index__"`. Expands `indexColumn` into the leading reserved column, which is what `exportExcel` does internally — the only public way to get the same result for the two low-level entry points above. Returns a **new** sheet (the input is not mutated; `merges` are shifted one column right) and throws if a user column already claims the reserved prop.
- `exportTable(options)` — `(options: TableExportOptions) => Promise<ExportResult>`. Convenience for table-shaped data. Full reference: [Table and ECharts entries](./05-table-and-echarts).
- `exportEcharts(options)` — `(options: EChartsExportOptions) => Promise<ExportResult>`. Convenience for ECharts data. Same reference page.

### Styling and formatting helpers

| Export                                                         | Signature                                                | Notes                                                                                                                                |
| -------------------------------------------------------------- | -------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| `StylePresets` / `StylePresetName`                             | `StylePresets` object; name is its key union             | 8 presets: `header` `currency` `percent` `date` `datetime` `dataRow` `bordered` `danger`                                             |
| `BaseCellStyle`                                                | `CellStyle` (frozen)                                     | The library-wide base style that every cell style is layered over (see the [Styles guide](/packages/excel-exporter/guide/05-styles)) |
| `applyFormat(value, spec)`                                     | `(value: unknown, spec: FormatSpec) => string \| number` | Applies a `FormatSpec` to a single value                                                                                             |
| `numFormatForSpec(spec)`                                       | `(spec: FormatSpec) => string \| null`                   | The Excel number-format code a `FormatSpec` resolves to                                                                              |
| `formatDateByPattern(value, pattern)`                          | `(value: unknown, pattern: string) => string`            | Date / datetime formatting under an explicit pattern                                                                                 |
| `displayValue(col, row)` / `resolveCellFormat(col, row)`       | `(col: ColumnConfig, row: Record<string, unknown>) => …` | What the exporter writes into a cell, and the format it resolves to — reusable in your own preview or pre-flight checks              |
| `validateSheetName(name)` / `validateMerges(sheet, leafCount)` | return `void`, throw on invalid input                    | The up-front validation `exportExcel` itself runs; call them to pre-check your input without exporting                               |
| `DEFAULT_DATE_PATTERN` / `DEFAULT_DATETIME_PATTERN`            | `"yyyy-MM-dd"` / `"yyyy-MM-dd HH:mm"`                    | The patterns `{ type: "date" }` / `{ type: "datetime" }` fall back to when no `pattern` is given                                     |

### WASM loader

`getWasmLoader()` returns this package's `WasmLoader` singleton. Observe it through the `isReady` / `supported` getters and `getOptions()` — the `LoadState` machine (`idle` / `loading` / `ready` / `error`) is private state, not a readable field. Both it and `configureWasm` [above](#configurewasm) are re-exports of the internal engine layer bundled into this package's `dist` (the same layer `@marcusok/excel-preview` bundles for itself — each package has its own loader instance).

### Overlay

`ProgressOverlayOptions` / `ProgressOverlayTextOptions` / `ProgressOverlayHandle` are re-exported so the `overlay` option can be typed without a second import. The overlay itself is bundled into this package; see the [Progress Overlay guide](/packages/excel-exporter/guide/11-overlay).

### Subpaths

| Subpath                                                                      | What it gives you                                                                                                                                                                                                                                                                                                                                                                      |
| ---------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `@marcusok/excel-exporter/styles`                                            | `StylePresets` + `StylePresetName` standalone — the same values as the main entry, for consumers that must not pull in the whole exporter                                                                                                                                                                                                                                              |
| `@marcusok/excel-exporter/worker-utils`                                      | `exportInWorker(options: ExportOptions, mode: "workbook" \| "stream"): Promise<ExportResult>` and `terminateWorker(): void`. **Not available from the main entry**, and `mode` is required. `terminateWorker()` terminates the shared worker and rejects every in-flight request with `worker terminated`                                                                              |
| `@marcusok/excel-exporter/overlay`                                           | Pre-2.8 compatibility layer: `exportExcelWithOverlay(options, overlay?)` and `showExportOverlay` (an alias of `showProgressOverlay` — the progress-overlay layer is bundled into the exporter and not published as a separate package). Mind the precedence — the wrapper's `overlay` parameter defaults to `{}`, so it **overrides** an `overlay: false` already present in `options` |
| `@marcusok/excel-exporter/dist/export.worker.js` · `…/dist/modern-xlsx.wasm` | The raw worker / WASM assets, for self-hosting when you do not want the loader's default URL resolution                                                                                                                                                                                                                                                                                |

```ts
import {
  exportExcel,
  configureWasm,
  WorkbookBuilder,
  exportAsStream,
  exportTable,
  exportEcharts,
  getWasmLoader,
} from "@marcusok/excel-exporter";
```

> This page covers the commonly used stable API. The main entry additionally re-exports the `format-utils` helpers tabulated above, `BorderStyle` and the other engine types from the bundled engine layer; `src/index.ts` is the exhaustive list.
