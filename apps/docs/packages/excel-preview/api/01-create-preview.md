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

| Option          | Type                                        | Default            | Description                                                                                                                                                                                                                   |
| --------------- | ------------------------------------------- | ------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `source`        | `File \| Blob \| Uint8Array \| ArrayBuffer` | — (required)       | The file bytes                                                                                                                                                                                                                |
| `password`      | `string`                                    | —                  | Password for encrypted workbooks (Agile AES-256)                                                                                                                                                                              |
| `sheet`         | `string \| number`                          | file's `activeTab` | Initial sheet (name or 0-based index). A missing name or an out-of-range index silently falls back to the file's `activeTab` — unlike the instance method `setSheet()`, which surfaces the invalid argument through `onError` |
| `showHeaders`   | `boolean`                                   | `true`             | Row/column headers (A/B/C + 1/2/3)                                                                                                                                                                                            |
| `showGridLines` | `boolean`                                   | from the file      | Grid lines                                                                                                                                                                                                                    |
| `showTabs`      | `boolean`                                   | `true`             | Sheet tab bar (hidden sheets never appear)                                                                                                                                                                                    |
| `onParsed`      | `(info: PreviewParsedInfo) => void`         | —                  | Render-ready callback: fires after the first parse+render **and again after every sheet switch** (`duration.parse` reuses the first parse's timing). Reports sheet list, dimensions and timings                               |
| `onError`       | `(error: PreviewError) => void`             | —                  | Failure callback (see codes below)                                                                                                                                                                                            |

`PreviewParsedInfo`:

```ts
interface PreviewParsedInfo {
  sheetNames: string[];
  sheetCount: number;
  rowCount: number; // active sheet
  colCount: number;
  duration: { parse: number; render: number; total: number }; // ms
}
```

## PreviewError codes

| Code                 | Meaning                                                                                                                                                               |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PASSWORD_PROTECTED` | Encrypted workbook, no/wrong `password` option                                                                                                                        |
| `LEGACY_FORMAT`      | Legacy `.xls` (BIFF8) — re-save as `.xlsx`                                                                                                                            |
| `CORRUPT`            | A ZIP that isn't a valid xlsx: missing/corrupt parts (`.ods` / `.docx`-style packages, truncated files) or a workbook with no sheets at all                           |
| `UNSUPPORTED`        | Not a ZIP and not recognizable plain text — includes XML/HTML "spreadsheets" (SpreadsheetML 2003, HTML tables saved as `.xls`) and UTF-16 CSV                         |
| `WASM`               | WebAssembly is unavailable in this environment, or the engine failed to load (404 asset URL, a CSP that forbids WebAssembly, network failure)                         |
| `UNKNOWN`            | Anything else; the underlying error message is passed through as-is (the original error object is in `.cause` when the failure came from the preview's own boot path) |

## PreviewInstance

```ts
interface PreviewInstance {
  destroy(): void; // unmount DOM, free resources
  setSheet(nameOrIndex: string | number): void; // switch sheets
  getSheetNames(): string[]; // all sheets in file order
}
```

`destroy()` removes the rendered DOM. The parse worker is a module-level shared resource (deliberately kept warm across instances, like the exporter's worker); it never holds file data after a parse completes.

## parseWorkbookBytes

The low-level parse, no DOM:

```ts
import { parseWorkbookBytes } from "@marcusok/excel-preview";

const workbook = await parseWorkbookBytes(bytes, { password: "…" });
```

Throws an `Error` with a `.code` property (same codes as above) on failure. See the [model types](/packages/excel-preview/api/02-model).

## Other exports

| Export            | What it is                                                                                            |
| ----------------- | ----------------------------------------------------------------------------------------------------- |
| `formatCellValue` | The cell formatter the renderer itself uses (see [model types](/packages/excel-preview/api/02-model)) |
| `configureWasm`   | Asset self-hosting config (see [assets](/packages/excel-preview/guide/02-assets))                     |
| `getWasmLoader`   | The shared loader singleton — current options/state (same object as `@marcusok/xlsx-core`'s)          |

`configureWasm` / `getWasmLoader` are re-exports of `@marcusok/xlsx-core`: importing them from either package configures the same shared loader.
