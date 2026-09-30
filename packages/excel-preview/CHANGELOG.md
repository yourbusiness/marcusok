# @marcusok/excel-preview

## 1.0.1

### Patch Changes

- 61fd7ba: Fix a set of style-restoration and robustness issues found in a full review:

  - Underlines are no longer lost on fonts written as `<u val="single"/>` (the explicit default form, common in third-party exports such as LibreOffice) — the val attribute is an enum, not a boolean.
  - Diagonal borders now respect the `diagonalUp`/`diagonalDown` flags: direction is rendered accordingly (`"/"` vs `"\"`), and a `<diagonal style>` without either flag no longer draws a line Excel itself would not show.
  - Gradient fills and diagonal borders no longer clobber each other — both land in one layered `background-image` declaration (diagonal on top).
  - Bold/italic are preserved when underline and strike are combined into one `text-decoration`.
  - A cross-origin `parseWorkerUrl` (raw CDN URL) now falls back to main-thread parsing with a console warning instead of throwing `SecurityError` out of the parse call; the docs now state that worker scripts are same-origin only.
  - A hung parse worker (WASM init stuck / engine loop) is terminated after `workerTimeoutMs` (default 120s), its in-flight requests are rejected without a main-thread rerun, and the next request rebuilds a fresh worker.

- Updated dependencies [61fd7ba]
  - @marcusok/xlsx-core@1.0.1

## 1.0.0

### Major Changes

- be12efc: Initial release: read-only xlsx preview in the browser. Worker-based parsing (shared worker, main-thread fallback in Node/SSR), framework-agnostic DOM rendering with four-quadrant frozen panes and virtual scrolling (viewport-only cells, whole-merge recall), and an Excel-fidelity layer: builtin number formats through Excel's actual behavior table, negative-sign/accounting-parens/currency-literal compensation, self-implemented dates/times (minute adjacency, weekday names, elapsed `[h]:mm:ss`, time-of-day `mm:ss` distinct from bracketed `[m]:ss`, fractional seconds, date1904 shift), General at 15 significant digits, and a theme/indexed color overlay that recovers colors the engine's read path drops (the overlay also rebuilds the fonts/fills/borders collections from styles.xml in document order — the engine's parser skips self-closed elements like the default `<border/>`, shifting every cellXfs id reference on real-Excel files). Supports encrypted workbooks (`password`), CSV (UTF-8/GB18030, delimiter sniffing), and friendly code-carrying errors (`PASSWORD_PROTECTED`, `LEGACY_FORMAT` for legacy .xls, …). Public API: `createPreview`, `parseWorkbookBytes`, `formatCellValue`.

### Patch Changes

- Updated dependencies [be12efc]
  - @marcusok/xlsx-core@1.0.0
