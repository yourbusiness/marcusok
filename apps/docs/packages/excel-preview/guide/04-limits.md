# Scope & Limits

What v1 deliberately does not do, and why.

## Not supported (friendly errors)

| Case                                                            | Behavior                                                                                                                                                                                |
| --------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Legacy `.xls` (BIFF8)                                           | `LEGACY_FORMAT` error with a "re-save as .xlsx" hint. OLE2/BIFF is a different binary world; a future optional adapter could bring SheetJS in, but v1 keeps the dependency chain clean. |
| `.ods` / other suites                                           | `UNSUPPORTED`/`CORRUPT` errors.                                                                                                                                                         |
| XML/HTML "spreadsheets" (SpreadsheetML 2003, HTML-table `.xls`) | `UNSUPPORTED` error — they are text but not CSV, so rendering them as a data grid is meaningless; re-save as a real `.xlsx`.                                                            |
| Corrupt / non-ZIP data                                          | `CORRUPT` error.                                                                                                                                                                        |
| Environments without WebAssembly                                | `WASM` error — xlsx parsing is WASM-only in v1 (there is no pure-JS fallback reader for xlsx); CSV still parses without WebAssembly.                                                    |

## Rendered with known approximations

- **Rich text** runs render as their concatenated plain text; run-level styling is on the roadmap (the current model keeps only the joined text, so restoring runs will need a parser addition).
- **Pattern fills** (hatched patterns like `gray125`) render as their background/foreground solid approximation; solid and gradient fills render exactly.
- **Number columns narrower than their content** clip instead of showing Excel's `####` (real `####` needs text measurement per cell — not worth the scroll cost yet).
- **External hyperlink URLs** are not rendered in v1; the hyperlink _styling_ (blue/underline from the file's font) still shows because it is ordinary font data. Tooltips are not surfaced either.
- **`rightToLeft` sheets** currently render left-to-right (the flag is parsed and carried in the model).
- **Conditional formatting** and data bars / icon sets are not applied.
- **Charts, images, shapes** are not rendered.
- **Row/column grouping** (outline levels) is neither parsed nor drawn.

## CSV encodings

UTF-8 (with or without BOM) and GB18030 are detected automatically. UTF-16
(`FF FE` / `FE FF` BOM) CSV files are detected as binary and rejected with
`UNSUPPORTED` — re-save them as UTF-8 first.

## Formulas

Formula cells render their **cached** `<v>` values — the same convention as SheetJS ("SheetJS does not evaluate formulas") and exceljs. A formula saved by a non-Excel writer without a cached value renders as an empty cell.

## Scale

Parsing is not streaming (OOXML parts cross-reference each other; browsers cannot stream-parse a zip of interdependent parts) — the whole file is read into memory inside the worker. 100k × 10 cells parse in ~1.5s (measured); virtual scrolling keeps the DOM at viewport size regardless of file dimensions. Multi-hundred-MB files will hit memory limits before anything else.
