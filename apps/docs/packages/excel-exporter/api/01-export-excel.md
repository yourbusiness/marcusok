# API: exportExcel & Configuration

## exportExcel

```ts
exportExcel(options: ExportOptions): Promise<ExportResult>
```

The core entry point (convenience wrappers such as `exportTable` / `exportEcharts` delegate to it). Routes to main / worker / stream by row count and environment, degrading to a style-less fast stream when WASM is unavailable. Auto-mode thresholds (browser): < 20,000 rows → main · 20,000–49,999 → worker + Workbook · ≥ 50,000 → worker + fast stream; Node/SSR: < 50,000 → main · ≥ 50,000 → stream — see [Auto Mode Routing](/packages/excel-exporter/guide/03-auto-mode).

## ExportOptions

| Field        | Type                                               | Required | Description                                                                                                                                                                                                                                                                                                                                                    |
| ------------ | -------------------------------------------------- | -------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `sheets`     | `SheetConfig[]`                                    | yes      | At least one sheet                                                                                                                                                                                                                                                                                                                                             |
| `filename`   | `string`                                           | yes      | Download name; `.xlsx` is appended unless it already ends with it (case-insensitive — `"REPORT.XLSX"` is left alone)                                                                                                                                                                                                                                           |
| `mode`       | `"auto" \| "main" \| "worker" \| "stream"`         | —        | Default `"auto"`. Validated: any other value (e.g. `"Main"`) fails up front with `{ success: false }` instead of silently routing as `auto`                                                                                                                                                                                                                    |
| `onProgress` | `(progress: number) => void`                       | —        | 0 → 1; the leading 0 and trailing 1 are each fired exactly once by `exportExcel` on every route (including the stream fallback and ultimately failed exports); incremental progress only on the stream path (every 1,000 rows)                                                                                                                                 |
| `onPhase`    | `(phase: ExportPhase, durationMs: number) => void` | —        | `init` / `build` / `download` timings, fired when each phase **ends** — see the reporting rules below                                                                                                                                                                                                                                                          |
| `download`   | `boolean`                                          | —        | Default `true`; `false` returns the Blob only                                                                                                                                                                                                                                                                                                                  |
| `overlay`    | `boolean \| ProgressOverlayOptions`                | —        | Default **on** since 2.8.0. `false` disables the full-screen progress overlay entirely; an options object customizes it — `delayMs`, `minVisibleMs`, `fadeOutMs`, `zIndex`, `container`, `blockInteraction`, `theme`, `text` (full field table in the guide). Node/SSR is a no-op. See the [Progress Overlay guide](/packages/excel-exporter/guide/11-overlay) |

Phase reporting rules (`onPhase`):

- **`init`** — WASM initialization. Main-thread routes measure `loader.ensureLoaded()`; the worker route reports it only when the worker actually (re-)initializes — a cached WASM instance reports nothing. The WASM-free stream routes on the main thread report a zero-duration `init` so the phase sequence stays stable; the browser worker-stream route (stream engine inside a worker) reports no `init` at all.
- **`build`** — workbook construction. Each **main-thread** build attempt reports its own `build` phase (in a `finally`, so also when that attempt throws — a degradation chain reports one per main-thread attempt). An attempt that fails _inside the worker_ reports no `build`; its failure surfaces only through the main-thread retry's phases and errors.
- **`download`** — the synchronous browser download trigger; reported only when `download !== false` **and** a `document` exists (never in Node).

## ExportResult

| Field       | Type            | Description                                                                                                                                                                                                                                                                                        |
| ----------- | --------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `success`   | `boolean`       | Whether the export succeeded                                                                                                                                                                                                                                                                       |
| `blob?`     | `Blob`          | The file content                                                                                                                                                                                                                                                                                   |
| `engine?`   | `"modern-xlsx"` | Engine actually used                                                                                                                                                                                                                                                                               |
| `mode?`     | `ExportMode`    | Mode actually used                                                                                                                                                                                                                                                                                 |
| `duration?` | `number`        | Total duration in ms. Main-thread routes measure the whole export; the worker route is measured on the main thread from the call into `exportInWorker` — including pre-post serialization and the worker round-trip, through Blob construction — so it is wider than the pure in-worker build time |
| `rowCount?` | `number`        | Data rows exported — header rows are not counted                                                                                                                                                                                                                                                   |
| `error?`    | `Error`         | Failure reason when `success` is `false`. **Also set — together with `success: true` —** when the export succeeded via the style-less stream fallback, carrying the degradation reason; check `success` first: a present `error` alone does not mean the export failed                             |

## configureWasm

```ts
configureWasm(options: LoaderOptions): void
```

Optional — assets default to the files shipped next to the package entry (see [Installation](/packages/excel-exporter/guide/02-installation)). Override for self-hosted copies, a CDN, or bundlers without asset-URL support.

| Field             | Default                 | Description                                                                                                                                                                                                       |
| ----------------- | ----------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `wasmUrl`         | the shipped `.wasm`     | Override for a self-hosted / CDN copy                                                                                                                                                                             |
| `workerUrl`       | the shipped worker file | Override for a **same-origin** self-hosted copy — a cross-origin URL is rejected at `Worker` construction, so a plain CDN does not work here                                                                      |
| `timeoutMs`       | `10_000`                | Per-attempt load timeout                                                                                                                                                                                          |
| `maxRetries`      | `3`                     | Max load attempts (3 total incl. the first; 300ms/600ms backoff)                                                                                                                                                  |
| `workerTimeoutMs` | `120_000`               | Worker export timeout. Keep it > 0 — there is no disable value: `0`/negative means "time out immediately", not "no timeout"                                                                                       |
| `parseWorkerUrl`  | —                       | Accepted by the shared `LoaderOptions` type but inert on this package: it points `@marcusok/excel-preview`'s parse worker at a self-hosted copy. This package's export worker is configured via `workerUrl` above |

## Other exported symbols

Everything below comes from the main entry. Grouped by what you reach for it for.

### Low-level build entries

- `WorkbookBuilder` — `static create(): Promise<WorkbookBuilder>` · `addSheet(config: SheetConfig): this` · `toBuffer(): Promise<Uint8Array>` · `toBlob(): Promise<Blob>`. Batch build with full styling, no mode routing. Requires WASM: `create()` initializes the engine, so in an environment without WebAssembly it fails outright — unlike `exportExcel`, this low-level entry has no stream fallback. Does **not** expand `SheetConfig.indexColumn` — pass an already-expanded sheet (see `applyIndexColumn` below), otherwise the field is silently ignored.
- `exportAsStream(sheets, onProgress?)` — `(sheets: SheetConfig[], onProgress?: (p: number) => void) => Promise<{ bytes: Uint8Array; rowCount: number }>`. The style-less fast stream, called directly. The second argument is the callback itself, **not** an options object. Same `indexColumn` caveat as `WorkbookBuilder`.
- `applyIndexColumn(sheet)` / `INDEX_PROP` — `(sheet: SheetConfig) => SheetConfig`, `INDEX_PROP = "__index__"`. Expands `indexColumn` into the leading reserved column, which is what `exportExcel` does internally — the only public way to get the same result for the two low-level entry points above. Returns a **new** sheet (the input is not mutated; `merges` are shifted one column right) and throws if a user column already claims the reserved prop.
- `exportTable(options)` — `(options: TableExportOptions) => Promise<ExportResult>`. Convenience for table-shaped data. Full reference: [Table and ECharts entries](./05-table-and-echarts).
- `exportEcharts(options)` — `(options: EChartsExportOptions) => Promise<ExportResult>`. Convenience for ECharts data. Same reference page.

### Styling and formatting helpers

| Export                                                         | Signature                                                | Notes                                                                                                                                                                                                                                         |
| -------------------------------------------------------------- | -------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `StylePresets` / `StylePresetName`                             | `StylePresets` object; name is its key union             | 8 presets: `header` `currency` `percent` `date` `datetime` `dataRow` `bordered` `danger`                                                                                                                                                      |
| `BaseCellStyle`                                                | `CellStyle` (frozen)                                     | The library-wide base style that every cell style is layered over (see the [Styles guide](/packages/excel-exporter/guide/05-styles))                                                                                                          |
| `applyFormat(value, spec)`                                     | `(value: unknown, spec: FormatSpec) => string \| number` | Applies a `FormatSpec` to a single value                                                                                                                                                                                                      |
| `toStr(value)`                                                 | `(value: unknown) => string`                             | The value stringifier the exporter itself uses for unformatted cells: objects → JSON (circular refs and odd `toJSON` shapes fall back to `String()`), `null`/`undefined` → `""`, Dates → ISO, bigints/booleans/symbols/functions → `String()` |
| `numFormatForSpec(spec)`                                       | `(spec: FormatSpec) => string \| null`                   | The Excel number-format code a `FormatSpec` resolves to                                                                                                                                                                                       |
| `formatDateByPattern(value, pattern)`                          | `(value: unknown, pattern: string) => string`            | Date / datetime formatting under an explicit pattern                                                                                                                                                                                          |
| `displayValue(col, row)` / `resolveCellFormat(col, row)`       | `(col: ColumnConfig, row: Record<string, unknown>) => …` | What the exporter writes into a cell, and the format it resolves to — reusable in your own preview or pre-flight checks                                                                                                                       |
| `validateSheetName(name)` / `validateMerges(sheet, leafCount)` | return `void`, throw on invalid input                    | The up-front validation `exportExcel` itself runs; call them to pre-check your input without exporting                                                                                                                                        |
| `DEFAULT_DATE_PATTERN` / `DEFAULT_DATETIME_PATTERN`            | `"yyyy-MM-dd"` / `"yyyy-MM-dd HH:mm"`                    | The patterns `{ type: "date" }` / `{ type: "datetime" }` fall back to when no `pattern` is given                                                                                                                                              |

### WASM loader

`getWasmLoader()` returns this package's `WasmLoader` singleton. Observe it through the `isReady` / `supported` getters and `getOptions()` — the `LoadState` machine (`idle` / `loading` / `ready` / `error`) is private state, not a readable field. Both it and `configureWasm` [above](#configurewasm) are re-exports of the internal engine layer bundled into this package's `dist` (the same layer `@marcusok/excel-preview` bundles for itself — each package has its own loader instance).

### Overlay

`ProgressOverlayOptions` / `ProgressOverlayTextOptions` / `ProgressOverlayHandle` are re-exported so the `overlay` option can be typed without a second import. The overlay itself is bundled into this package; see the [Progress Overlay guide](/packages/excel-exporter/guide/11-overlay).

### Subpaths

| Subpath                                                                      | What it gives you                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| ---------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `@marcusok/excel-exporter/styles`                                            | `StylePresets` + `StylePresetName` standalone — the same values as the main entry, for consumers that must not pull in the whole exporter                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `@marcusok/excel-exporter/worker-utils`                                      | `exportInWorker(options: ExportOptions, mode: "workbook" \| "stream"): Promise<ExportResult>` and `terminateWorker(): void`. **Not available from the main entry**, and `mode` is required. `terminateWorker()` terminates the shared worker and rejects every in-flight request with `worker terminated`                                                                                                                                                                                                                                                                                                |
| `@marcusok/excel-exporter/overlay`                                           | Pre-2.8 compatibility layer plus the generic overlay entry: `exportExcelWithOverlay(options, overlay?)`, `showExportOverlay` (an alias of `showProgressOverlay` — the progress-overlay layer is bundled into the exporter and not published as a separate package), `nextPaint()` (the only public entry point of that helper — the main entry does not export it) and the renamed types `ExportOverlayHandle` / `OverlayOptions` / `OverlayTextOptions`. Mind the precedence — the wrapper's `overlay` parameter defaults to `{}`, so it **overrides** an `overlay: false` already present in `options` |
| `@marcusok/excel-exporter/dist/export.worker.js` · `…/dist/modern-xlsx.wasm` | The raw worker / WASM assets, for self-hosting when you do not want the loader's default URL resolution                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |

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

> This page covers the commonly used stable API. The main entry additionally re-exports the `format-utils` helpers tabulated above, `BorderStyle` (declared in this package's own `types.ts`) and the engine-layer types `LoaderOptions` / `LoadState` (importable as types from the main entry); the exhaustive export list is [`src/index.ts` on GitHub](https://github.com/yourbusiness/marcusok/blob/main/packages/excel-exporter/src/index.ts) (shipped in the installed package as `dist/index.d.ts`).
