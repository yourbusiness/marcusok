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

| Option          | Type                                        | Default            | Description                                                                                                                                                                                                                                                                                                                                         |
| --------------- | ------------------------------------------- | ------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `source`        | `File \| Blob \| Uint8Array \| ArrayBuffer` | — (required)       | The file bytes                                                                                                                                                                                                                                                                                                                                      |
| `password`      | `string`                                    | —                  | Password for encrypted workbooks (Agile AES-256)                                                                                                                                                                                                                                                                                                    |
| `sheet`         | `string \| number`                          | file's `activeTab` | Initial sheet (name or 0-based index). An invalid value (unknown name, out-of-range or non-integer index) reports `Sheet not found: …` through `onError` — same wording and code as the instance method `setSheet()` — and still falls back to rendering the file's `activeTab` (itself normalized — see `activeSheetIndex` in the model reference) |
| `showHeaders`   | `boolean`                                   | `true`             | Row/column headers (A/B/C + 1/2/3)                                                                                                                                                                                                                                                                                                                  |
| `showGridLines` | `boolean`                                   | from the file      | Grid lines                                                                                                                                                                                                                                                                                                                                          |
| `showTabs`      | `boolean`                                   | `true`             | Sheet tab bar (hidden sheets never appear)                                                                                                                                                                                                                                                                                                          |
| `styleNonce`    | `string`                                    | —                  | CSP nonce applied to the injected `<style>` elements. Cell geometry still uses inline `style` attributes, so a strict CSP must also allow inline style attributes                                                                                                                                                                                   |
| `onParsed`      | `(info: PreviewParsedInfo) => void`         | —                  | Render-ready callback: fires after the first parse+render **and again after every sheet switch** (`duration.parse` reuses the first parse's timing, `duration.total` keeps counting from the initial load — for a switch's own cost use `duration.render`). Reports sheet list, dimensions and timings                                              |
| `onError`       | `(error: PreviewError) => void`             | —                  | Failure callback (see codes below)                                                                                                                                                                                                                                                                                                                  |

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

| Code                 | Meaning                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PASSWORD_PROTECTED` | Encrypted workbook, no/wrong `password` option                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| `LEGACY_FORMAT`      | Legacy `.xls` (BIFF8) — re-save as `.xlsx`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `CORRUPT`            | A ZIP that isn't a valid xlsx: missing/corrupt parts (`.ods` / `.docx`-style packages, truncated files) or a workbook with no sheets at all                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| `UNSUPPORTED`        | Not a ZIP and not recognizable plain text — includes XML/HTML "spreadsheets" (SpreadsheetML 2003, HTML tables saved as `.xls`) and UTF-16 CSV                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| `WASM`               | WebAssembly is unavailable in this environment, or the engine failed to load (404 asset URL, a CSP that forbids WebAssembly, network failure)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| `UNKNOWN`            | Anything else — including a parse-worker timeout, which carries the dedicated message `parse worker timed out after …ms (worker terminated); retry the parse` (the worker is terminated and the request is **not** re-run on the main thread); the underlying error message is passed through as-is (the original error object is in `.cause` when the failure came from the preview's own boot path)                                                                                                                                                                                                                                                                                                               |
| `ABORTED`            | The caller cancelled the parse — `destroy()` on an in-flight preview, or an aborted `signal` passed to `parseWorkbookBytes`. Cancellation is never routed to the main-thread fallback (re-running a cancelled input would defeat the cancellation). `createPreview` treats it silently: `onError` is **not** fired for it. Semantics: on the main-thread path the parse gives up at the next `await` boundary (the synchronous WASM section cannot be interrupted mid-flight); on the worker path the caller stops waiting immediately and the late worker result is dropped via its already-cleared pending slot — the shared worker is **not** terminated for a cancellation (that would reject sibling requests) |

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

`destroy()` removes the rendered DOM **and aborts an in-flight parse** (the `ABORTED` code above — the boot path returns silently, `onError`/`onParsed` never fire for a cancelled preview). The parse worker is a module-level shared resource (deliberately kept warm across instances, like the exporter's worker); it never holds file data after a parse completes. Worker requests are serialized through a main-thread FIFO queue, so aborting the active parse terminates the worker (queued parses are re-dispatched to a fresh instance); aborting a queued parse only removes that request.

Hidden-sheet semantics: `setSheet()` looks the sheet up by name or 0-based index and **does not filter hidden sheets** — unlike the tab bar (which never shows them), an explicit `setSheet("HiddenSheet")` renders it. An invalid name/index (including non-integer numbers and `NaN`) surfaces through `onError` (`UNKNOWN` code, message `Sheet not found: …`) instead of throwing; `activeSheetIndex` keeps the previous valid sheet. `getSheetNames()` likewise includes hidden / veryHidden sheets, in file order — filter on your side if you drive your own sheet picker from it.

User callbacks are isolated: a throw inside `onParsed` / `onError` is reported with a `console.warn` and cannot be mistaken for a parse/render failure or produce an unhandled rejection.

## parseWorkbookBytes

The low-level parse, no DOM:

```ts
import { parseWorkbookBytes } from "@marcusok/excel-preview";

const workbook = await parseWorkbookBytes(bytes, {
  password: "…",
  signal, // optional AbortSignal — abort rejects with an ABORTED-coded error
});
```

Throws an `Error` with a `.code` property (same codes as above) on failure; an aborted `signal` rejects with code `ABORTED`. See the [model types](/packages/excel-preview/api/02-model).

## Other exports

| Export                 | What it is                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| ---------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `formatCellValue`      | The cell formatter the renderer itself uses; returns a `FormattedValue` (`{ text, color? }`, also exported as a type — see [model types](/packages/excel-preview/api/02-model))                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `configureWasm`        | Asset self-hosting config (see [assets](/packages/excel-preview/guide/02-assets))                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| `getWasmLoader`        | This package's `WasmLoader` singleton (the engine layer bundled into its `dist`). Read readiness via `isReady` / `supported`, options via `getOptions()` (the state machine itself is private). `ensureLoaded()` is public as well — but pre-warming only pays off on the main-thread routes (Node/SSR, or a session that fell back after a worker construct/load failure). In the normal browser route the parse runs inside the worker, which loads its own WASM copy — pre-warming the main thread just downloads a second 1.9MB binary for nothing, and there is no public API to spin the worker up early |
| `terminateParseWorker` | Terminates the shared parse worker and rejects its in-flight **and queued** parses (`noFallback` — no main-thread rerun, same policy as a worker timeout). This is the escape hatch after changing `parseWorkerUrl`/`wasmUrl`: the worker reads its URLs once at creation, and the next parse rebuilds a fresh worker from the current options. It also clears a previous worker construction/load failure (`workerBroken`), which otherwise keeps the session on the main-thread path until reload. Recreating a preview is not enough — the worker is a module-level singleton shared across instances       |

`configureWasm` / `getWasmLoader` come from the engine layer bundled into this package: each @marcusok business package carries its own copy and its own loader instance — configure the one you use.
