# createPreview

The high-level entry: parse + render in one call.

```ts
import { createPreview } from "@marcusok/excel-preview";

function createPreview(
  container: HTMLElement,
  options: PreviewOptions,
): PreviewInstance;
```

## PreviewOptions

| Option          | Type                                        | Default            | Description                                                                                                                                                                                                                                                                                            |
| --------------- | ------------------------------------------- | ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `source`        | `File \| Blob \| Uint8Array \| ArrayBuffer` | — (required)       | The file bytes                                                                                                                                                                                                                                                                                         |
| `password`      | `string`                                    | —                  | Password for encrypted workbooks (Agile AES-256)                                                                                                                                                                                                                                                       |
| `sheet`         | `string \| number`                          | file's `activeTab` | Initial sheet (name or 0-based index). A missing name or an out-of-range index silently falls back to the file's `activeTab` (itself normalized — see `activeSheetIndex` in the model reference) — unlike the instance method `setSheet()`, which surfaces the invalid argument through `onError`      |
| `showHeaders`   | `boolean`                                   | `true`             | Row/column headers (A/B/C + 1/2/3)                                                                                                                                                                                                                                                                     |
| `showGridLines` | `boolean`                                   | from the file      | Grid lines                                                                                                                                                                                                                                                                                             |
| `showTabs`      | `boolean`                                   | `true`             | Sheet tab bar (hidden sheets never appear)                                                                                                                                                                                                                                                             |
| `onParsed`      | `(info: PreviewParsedInfo) => void`         | —                  | Render-ready callback: fires after the first parse+render **and again after every sheet switch** (`duration.parse` reuses the first parse's timing, `duration.total` keeps counting from the initial load — for a switch's own cost use `duration.render`). Reports sheet list, dimensions and timings |
| `onError`       | `(error: PreviewError) => void`             | —                  | Failure callback (see codes below)                                                                                                                                                                                                                                                                     |

`PreviewParsedInfo`:

```ts
interface PreviewParsedInfo {
  sheetNames: string[];
  sheetCount: number;
  rowCount: number; // active sheet
  colCount: number;
  duration: {
    parse: number; // ms; the first parse only — reused on later sheet switches
    render: number; // ms; this render
    total: number; // ms; elapsed since createPreview booted — cumulative, not per-switch
  };
}
```

## PreviewError codes

| Code                 | Meaning                                                                                                                                                                                                                                                                                                                                                                                               |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PASSWORD_PROTECTED` | Encrypted workbook, no/wrong `password` option                                                                                                                                                                                                                                                                                                                                                        |
| `LEGACY_FORMAT`      | Legacy `.xls` (BIFF8) — re-save as `.xlsx`                                                                                                                                                                                                                                                                                                                                                            |
| `CORRUPT`            | A ZIP that isn't a valid xlsx: missing/corrupt parts (`.ods` / `.docx`-style packages, truncated files) or a workbook with no sheets at all                                                                                                                                                                                                                                                           |
| `UNSUPPORTED`        | Not a ZIP and not recognizable plain text — includes XML/HTML "spreadsheets" (SpreadsheetML 2003, HTML tables saved as `.xls`) and UTF-16 CSV                                                                                                                                                                                                                                                         |
| `WASM`               | WebAssembly is unavailable in this environment, or the engine failed to load (404 asset URL, a CSP that forbids WebAssembly, network failure)                                                                                                                                                                                                                                                         |
| `UNKNOWN`            | Anything else — including a parse-worker timeout, which carries the dedicated message `parse worker timed out after …ms (worker terminated); retry the parse` (the worker is terminated and the request is **not** re-run on the main thread); the underlying error message is passed through as-is (the original error object is in `.cause` when the failure came from the preview's own boot path) |

The error object's full shape:

```ts
interface PreviewError {
  code: PreviewErrorCode;
  message: string;
  cause?: unknown; // the original error, boot-path exceptions only
}
```

`cause` is attached only by the boot-path exception handler; parse failures surfaced through the parse pipeline carry `code` + `message` without a `cause` — don't build diagnostics on it always being there.

## PreviewInstance

```ts
interface PreviewInstance {
  destroy(): void; // unmount DOM, free resources
  setSheet(nameOrIndex: string | number): void; // switch sheets
  getSheetNames(): string[]; // all sheets in file order, hidden ones included
}
```

`destroy()` removes the rendered DOM. The parse worker is a module-level shared resource (deliberately kept warm across instances, like the exporter's worker); it never holds file data after a parse completes.

Hidden-sheet semantics: `setSheet()` looks the sheet up by name or 0-based index and **does not filter hidden sheets** — unlike the tab bar (which never shows them), an explicit `setSheet("HiddenSheet")` renders it. An invalid name/index surfaces through `onError` (`UNKNOWN` code). `getSheetNames()` likewise includes hidden / veryHidden sheets, in file order — filter on your side if you drive your own sheet picker from it.

## parseWorkbookBytes

The low-level parse, no DOM:

```ts
import { parseWorkbookBytes } from "@marcusok/excel-preview";

const workbook = await parseWorkbookBytes(bytes, { password: "…" });
```

Throws an `Error` with a `.code` property (same codes as above) on failure. See the [model types](/packages/excel-preview/api/02-model).

## Other exports

| Export            | What it is                                                                                                                                                                                                                                                                                                                                |
| ----------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `formatCellValue` | The cell formatter the renderer itself uses; returns a `FormattedValue` (`{ text, color? }`, also exported as a type — see [model types](/packages/excel-preview/api/02-model))                                                                                                                                                           |
| `configureWasm`   | Asset self-hosting config (see [assets](/packages/excel-preview/guide/02-assets))                                                                                                                                                                                                                                                         |
| `getWasmLoader`   | This package's `WasmLoader` singleton (the engine layer bundled into its `dist`). Read readiness via `isReady` / `supported`, options via `getOptions()` (the state machine itself is private). `ensureLoaded()` is public as well — call it ahead of the first preview to pre-warm the engine and take the WASM load off the first parse |

`configureWasm` / `getWasmLoader` come from the engine layer bundled into this package: each @marcusok business package carries its own copy and its own loader instance — configure the one you use.
