# Scope & Limits

What the package deliberately does not do, and why.

## Not supported (friendly errors)

| Case                                                            | Behavior                                                                                                                                                                                                                                                                                                                                                                               |
| --------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Legacy `.xls` (BIFF8)                                           | `LEGACY_FORMAT` error with a "re-save as .xlsx" hint. OLE2/BIFF is a different binary world; a future optional adapter could bring SheetJS in, but the package keeps the dependency chain clean.                                                                                                                                                                                       |
| `.ods` / other suites (also ZIP packages)                       | `CORRUPT` error — they are ZIP archives without an `xl/workbook.xml`, which is reported as "not a valid xlsx".                                                                                                                                                                                                                                                                         |
| XML/HTML "spreadsheets" (SpreadsheetML 2003, HTML-table `.xls`) | `UNSUPPORTED` error — they are text but not CSV, so rendering them as a data grid is meaningless; re-save as a real `.xlsx`.                                                                                                                                                                                                                                                           |
| Corrupt data                                                    | Broken ZIP structure (truncated files, …) or a workbook with no sheets at all → `CORRUPT`; not a ZIP and no readable text either → `UNSUPPORTED` (the format cannot be identified).                                                                                                                                                                                                    |
| Out-of-range declared dimensions                                | `CORRUPT` — a row index above 1,048,576, a column above XFD (16,384), or a merge / `<col>` span beyond those limits. The check runs during model build (inside the worker), before the main-thread layout allocates geometry arrays, so a tiny crafted file cannot force multi-GB allocation. Negative/non-finite row heights and column widths are normalized to the default instead. |
| Environments without WebAssembly / engine load failure          | `WASM` error — xlsx parsing is WASM-only (there is no pure-JS fallback reader for xlsx); a 404 asset URL, a CSP that forbids WebAssembly or a network failure reports `WASM` too. CSV still parses without WebAssembly.                                                                                                                                                                |

## Rendered with known approximations

- **Rich text** runs render as their concatenated plain text; run-level styling is on the roadmap (the current model keeps only the joined text, so restoring runs will need a parser addition).
- **Pattern fills** (hatched patterns like `gray125`) render as their background/foreground solid approximation; solid and gradient fills render exactly.
- **Rotated text** (`textRotation`) is clipped to the cell box when it does not fit — Excel also truncates rotated text that exceeds the cell, but its clipping follows the rotated text run rather than the axis-aligned cell rectangle, so edge cases near the diagonal differ slightly.
- **Number columns narrower than their content** clip instead of showing Excel's `####` (real `####` needs text measurement per cell — not worth the scroll cost yet).
- **Oversized merges** skip the per-cell coverage expansion when a single merge exceeds 1,000,000 covered cells (`rowSpan × colSpan`), or when the expansions of all merges combined exceed 2,000,000 cells (budget consumed in file order): the merge's anchor cell still renders with its span, but covered cells are no longer blanked out, so stray values stored inside such a merge become visible (Excel semantics would show them as empty). The caps exist because the expansion runs on the main thread — a whole-sheet merge, or many near-limit merges stacked, would freeze the page.
- **Initial model handoff** runs on the main thread: the worker parses off-thread, but the finished `PreviewWorkbook` is structured-cloned back, which is one main-thread task that grows with the cell count (measured ~50ms at 10k rows, ~150ms at 50k and ~300ms at 100k in Chromium). Scrolling and rendering stay virtualized; expect a brief pause at the end of the parse, not during it.
- **Node runtime sensitivity**: on Node 22.x (and Node 24.14), the bundled engine's `readBuffer` is ~10x slower than on Node 24.21 for numeric-heavy sheets. The package still works, but the browser/pinned-runtime performance numbers do not apply. See the repository README for the support matrix.
- **Split panes** (`splitPane`) are ignored — only freeze panes render (the model carries `frozenPane` only).
- **Tab colors** are not carried (the tab bar renders without the sheet's tab color); cell comments are dropped by the read path as well.
- **External hyperlink URLs** are not rendered; the hyperlink _styling_ (blue/underline from the file's font) still shows because it is ordinary font data. Tooltips are not surfaced either.
- **`rightToLeft` sheets** currently render left-to-right (the flag is parsed and carried in the model).
- **Conditional formatting** and data bars / icon sets are not applied.
- **Charts, images, shapes** are not rendered.
- **Row/column grouping** (outline levels) is neither parsed nor drawn.

## CSV encodings

UTF-8 (with or without BOM) and GB18030 are detected automatically. UTF-16
CSV files usually fail the text sniff — not because their BOM (`FF FE` /
`FE FF`) is detected, but because UTF-16 bytes are full of NULs and the sniff
requires ≥ 95% printable bytes in the first 2 KB. They land in `unknown` and
are rejected with `UNSUPPORTED` — re-save them as UTF-8 first (a
pathologically printable UTF-16 stream could still slip past the sniff and
decode as garbage instead of being rejected).

## Formulas

Formula cells render their **cached** `<v>` values — the same convention as SheetJS ("SheetJS does not evaluate formulas") and exceljs. A formula saved by a non-Excel writer without a cached value renders as an empty cell.

## Scale

Parsing is not streaming (OOXML parts cross-reference each other; browsers cannot stream-parse a zip of interdependent parts) — the whole file is read into memory inside the worker. 100k rows × 10 cols parse in ~1.5s (measured); virtual scrolling keeps the DOM at viewport size regardless of file dimensions. Multi-hundred-MB files will hit memory limits before anything else.
The repo's performance test asserts a different, looser bound: 100k rows × 8 cols parse + model build < 4s, measured inside happy-dom without real rendering, and skipped in CI and the release pipeline (`RUN_PERF=0`).

**Row-count ceiling (a browser limit, not memory)**: the grid is one tall container element, and at 20px per row (15pt) roughly 890k rows (~17.9M px) hit the element-size cap in Firefox / Safari — rows past that point are unreachable. Chrome's cap is ~33.5M px (about 1.67M rows). Headers and row/column labels stay viewport-sized either way; only the scrollable range is cut short. A full 1,048,576-row sheet sits between the two caps (~21M px at 20px rows — reachable in Chrome, truncated in Firefox / Safari), and exceeds every browser once rows are taller.
