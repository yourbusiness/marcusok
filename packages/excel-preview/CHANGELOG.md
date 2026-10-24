# @marcusok/excel-preview

## 2.4.2

### Patch Changes

- b23fd99: Fix custom number formats with quoted literals (e.g. `yyyy"年"m"月"d"日"`, `0"元"`) rendering as `&quot;` garbled text. The engine returns the raw `styles.xml` attribute value without XML entity decoding; the model layer now decodes `formatCode` before resolving it into `PreviewXf.numFmtCode`.

## 2.4.1

### Patch Changes

- 651fbde: Make the published `.d.ts` fully self-contained by inlining the two engine type references (`BorderStyle`, `AlignmentData`, now also exported for custom renderers), and move `modern-xlsx` back to devDependencies. Consumers no longer install `modern-xlsx` as a transitive runtime dependency, which also drops its `engines.node >= 24` declaration from consumer installs (the runtime engine was already bundled into dist; nothing changes at runtime).

## 2.4.0

### Minor Changes

- aba03fe: Harden export input handling, routing and preview dimensions.

  - `excel-exporter`: reject XML 1.0-illegal characters in sheet names, font names, number formats and date/datetime patterns (they previously produced workbooks Excel flags as corrupt while the export still returned `success: true`); sanitize data/header text consistently on both routes; runtime-validate style enums and types for JS callers; include the estimated cell count in auto routing so wide sheets no longer take the slow Workbook path; expose dropped fast-stream features as `ExportResult.degradedFeatures`; inject overlay styles/nodes into the container's own document (iframe support).
  - `excel-preview`: reject out-of-range row/column/merge dimensions as `CORRUPT` before the main-thread layout allocates geometry arrays; normalize negative/non-finite row heights and column widths; inject renderer styles and create DOM nodes in the container's own document (iframe support).

## 2.3.0

### Minor Changes

- 25093cd: Improve embedding and worker robustness: styles are injected into the container's ShadowRoot (with an optional CSP `styleNonce`) and the viewport window is defensively capped when CSS did not apply; the parse worker now uses a FIFO queue where timeouts measure processing time and do not reject queued parses; `terminateParseWorker()` clears the `workerBroken` fallback; non-integer/NaN `setSheet()` indexes report through `onError` instead of throwing; and `onParsed`/`onError` callback exceptions are isolated.

### Patch Changes

- d883ceb: Support local WASM files in Node: `configureWasm({ wasmUrl })` now accepts a plain filesystem path or a `file:` URL in addition to HTTP(S) URLs. The binary is read synchronously (matching zero-config auto-init), and a missing/unreadable configured file fails with the real fs/compile reason instead of a misleading `fetch failed`.

## 2.2.0

### Minor Changes

- 210d22b: Add `terminateParseWorker()` — a public escape hatch that terminates the shared parse worker and rejects its in-flight parses (`noFallback`, no main-thread rerun; same policy as a worker timeout). Changing `parseWorkerUrl`/`wasmUrl` previously had no public way to recycle the already-created worker: the worker is a module-level singleton, and the shared loader's guidance suggested recreating the preview, which does not recreate the worker. After `terminateParseWorker()`, the next parse rebuilds a fresh worker from the current configuration. The shared loader's reconfiguration warnings (bundled from xlsx-core) now point at this escape hatch instead of the ineffective advice.

## 2.1.4

### Patch Changes

- 9226572: Cap the total number of merge-covered cells expanded during layout at 2,000,000 cells across all merges, on top of the existing per-merge 1,000,000 cap. Files declaring many near-limit merges could previously stack the main-thread expansion into multi-second freezes (10 × 1M merges ≈ 5.5s measured); over-budget merges now degrade exactly like single oversized ones — the anchor index is kept so the merge still renders with its span, while covered cells fall back to rendering their own content.

## 2.1.3

### Patch Changes

- d69d0b0: Fix merged cells disappearing from frozen bands: the frozen row/column bands follow the main viewport on their unpinned axis, so a wide merge anchored inside the band (e.g. an `A1:H1` title in a frozen first row) vanished once scrolling moved its anchor out of the window. Band layers now recall off-window merge anchors the same way the main quadrant does.
- f32c23c: Cap the per-merge cell expansion in layout building: a pathological whole-sheet merge (`A1:XFD1048576`, ~1.7e10 cells) previously froze the main thread while populating the covered-cell index. Merges over 1M cells now keep only their anchor entry (the anchor still renders at full span); normal large merges such as full-column merges are unaffected.

## 2.1.2

### Patch Changes

- 03b2a1a: Tightened the text-format sniff: a mostly-non-ASCII binary whose printable ratio passes the 95% heuristic (e.g. a 0xFF-filled blob) is no longer rendered as garbled CSV — the sampled bytes must now decode as UTF-8 or GB18030, otherwise the parse fails with `UNSUPPORTED`. GBK/GB18030 CSVs, including ones truncated mid-character at the 2KB sniff window, still parse as before. Also synced the stale `PreviewOptions.sheet` JSDoc with the actual behavior (an invalid initial sheet reports `Sheet not found` through `onError` before falling back to `activeTab`).

## 2.1.1

### Patch Changes

- 43ba9ed: An invalid initial `sheet` option (unknown name, out-of-range or non-integer index) now reports `Sheet not found: …` through `onError` — same wording and code as the instance method `setSheet()` — instead of silently falling back. The fallback rendering behavior is unchanged: the preview still renders the file's `activeTab`.

## 2.1.0

### Minor Changes

- 3b058bc: Add parse cancellation. `createPreview().destroy()` now aborts an in-flight parse immediately instead of letting it run to completion: the boot path returns silently (`onError` is never fired for a cancellation), the main-thread fallback path gives up at the next `await` boundary, and a late worker result is dropped via its already-cleared pending slot (the shared worker is not terminated — that would reject sibling requests). The low-level `parseWorkbookBytes(bytes, opts)` accepts an optional `AbortSignal` via `opts.signal` and rejects with an `ABORTED`-coded error when aborted. `ABORTED` is a new `PreviewErrorCode` member and is never routed to the main-thread fallback.

## 2.0.8

### Patch Changes

- f472648: Docs: describe the main-thread fallback precisely. The README now states both trigger conditions (no `Worker` or no `window` global) and that a worker which fails to construct/load falls back permanently for the session, instead of attributing the fallback to "no Worker global" (Node/SSR) alone.

## 2.0.7

### Patch Changes

- b971c17: Include CHANGELOG.md in the published npm package (files allowlist), matching @marcusok/excel-exporter so the changelog is readable on npm alongside the README.

## 2.0.6

### Patch Changes

- caed971: Harden the styles-overlay XML entity decoder against malicious character references: an out-of-range numeric entity (`&#1114112;`) previously threw a `RangeError` that escaped the overlay's degradation guard and failed the whole parse with an `UNKNOWN` error, instead of falling back to the engine style arrays. Such entities (and lone surrogates) are now replaced with U+FFFD, and the hexadecimal character-reference form (`&#x41;`) is decoded per the XML spec.

## 2.0.5

### Patch Changes

- 6d4ec14: Align two edge behaviors with Excel semantics: text overflow is no longer blocked by value-less styled cells (an empty cell blocks spill in Excel regardless of formatting, so the spill limit now extends to the next cell with content or a merged region), and CSV values with a trailing decimal point (`"123."`) are now inferred as numbers, matching Excel's input-bar interpretation (leading-dot `".5"` and multi-dot `"1.2.3"` forms stay text).

## 2.0.4

### Patch Changes

- 609ddb5: Fix a ResizeObserver leak on repeated `render()` calls (e.g. loading another file into the same renderer instance). The previous observer — still watching the detached viewport — is now disconnected before a new one is created, so observers no longer accumulate until `destroy()`.

## 2.0.3

### Patch Changes

- 4eb4d3b: Fix number-format rendering for fractions and scientific notation: the display-rounding pre-pass no longer pre-rounds fraction cores (`# ?/?`, `# ??/??` — previously `1.25` rendered as `1`) or scientific cores (`0.00E+00` — previously tiny values like `0.00012` collapsed to `0.00E+0`), and scientific exponents are normalized to Excel's two-digit signed form (`1.23E+4` → `1.23E+04`).

## 2.0.2

### Patch Changes

- 8c7a5a4: Number-format fidelity fixes for zh-CN Excel files and Excel-exact rounding:

  - Builtin `numFmtId` 27–36 and 50–58 (East-Asian groups, only written by East-Asian-locale Excel/WPS) now resolve to zh-CN behavior strings — zh-CN dates stored as builtin ids (e.g. id 31 for `2025年1月21日`) render as dates instead of raw serials like `45678`.
  - `[DBNum1/2/3]` digit-substitution modifiers are stripped so the numeric section still renders; unknown bracket keywords no longer misread the section as a date (`[DBNum2][$-804]#,##0.00` used to render as `[18BNu52]#,##0.00`).
  - Display rounding now follows Excel's decimal semantics (15 significant digits, half-up): `2.675` with `0.00` shows `2.68` (the engine's binary `toFixed` produced `2.67`); percentages round in the display domain (`0.12345` with `0.00%` → `12.35%`).
  - Trailing-comma thousands scaling is supported (`#,##0,` divides by 1000 per trailing comma, `0.0,,"M"` → `1.5M`).
  - Section-level `General` (`[Red]General`, negative-section `General;-General`, WPS's `G/通用格式`) renders the value instead of the literal text "General".

## 2.0.1

### Patch Changes

- 1ed9f84: Fix formatting and alignment fidelity issues:

  - Cells with an explicit `horizontal="general"` alignment now fall back to type-based alignment (numbers render right-aligned), matching Excel semantics for the default alignment.
  - A workbook whose `activeTab` points at a hidden sheet now opens the first visible sheet instead of silently rendering hidden content with no tab to switch away.
  - Literal-only number formats (e.g. `"yes"`) no longer append the numeric value — Excel hides the number when a section has no digit placeholder.
  - Number formats whose sections are all conditional with no match (e.g. `[>100]0.00` for 5) now render empty instead of falling back to the first conditional section.
  - Font names are sanitized more strictly (control characters stripped) when compiling the built-in stylesheet, closing a CSS-injection vector from untrusted xlsx files.

## 2.0.0

### Major Changes

- 6a3729c: Bundle the engine layer into the package and drop the same-scope runtime dependency. `@marcusok/xlsx-core` is no longer published to npm; it is a private build-time layer now bundled into this package's `dist`, so installing `@marcusok/excel-preview` brings the whole parsing engine and the WASM binary. Migration notes:

  - **Remove any direct `@marcusok/xlsx-core` imports** — the package is no longer installable. `configureWasm` / `getWasmLoader` were already re-exported from the main entry.
  - **WASM asset path**: the canonical self-hosting path is now `@marcusok/excel-preview/dist/modern-xlsx.wasm?url` (previously `@marcusok/xlsx-core/dist/modern-xlsx.wasm`).
  - **One loader per package**: `configureWasm` from this package configures this package's bundled loader only; `@marcusok/excel-exporter` has its own. On a page using both, call it once per package (passing the same `wasmUrl` is fine — the binaries are identical).
  - The only remaining `dependencies` entry is a types-only pinned `modern-xlsx` (published `.d.ts` files keep external type imports resolvable); nothing from it loads at runtime.

## 1.0.3

### Patch Changes

- 020bb13: Fix three preview defects. (1) Sheet-level CSS is now scoped per renderer instance, so mounting several previews on one page no longer cross-wires them: previously the cell-style rules (`.xpv-xf-N`) and the grid-line rule were global, so the last-mounted instance's rules won the cascade for every instance — styles from one file leaked into another, and an instance with `showGridLines: false` still drew grid lines when another instance had them on. (2) Error codes now match the documented contract: a ZIP that isn't a valid xlsx (`.ods` / `.docx`-style packages, truncated files) and a workbook with no sheets report `CORRUPT` instead of `UNKNOWN`, and a failed WebAssembly load reports `WASM` instead of `UNKNOWN`; `UNKNOWN` is now reserved for genuinely unclassified failures. A workbook with zero sheets also surfaces as an error instead of rendering an empty preview with no `onParsed` / `onError` callback. (3) Formula cells whose cached result is a string now overflow into adjacent empty cells, matching literal text (Excel spills both).

## 1.0.2

### Patch Changes

- a81b18a: Fix a set of rendering and number-format issues found in a second review:

  - Uppercase date/time format codes no longer misread `M` as a month: the minute/ month adjacency check was case-sensitive while format tokens are case-insensitive, so third-party codes like `HH:MM:SS` rendered as hour:month:second (e.g. `06:12:00` instead of `06:00:00`).
  - Huge merged cells no longer break virtual scrolling. Merge recall used to widen the visible window to the whole merge rectangle, so a full-column merge (A1:A100000 — common in templates) made the renderer build every row header in the merge (100k+ DOM nodes). Only the merge's anchor cell is recalled now; headers stay at viewport size.
  - Text spill is capped at the next non-empty cell in the row (Excel's truncation semantics) — overflowing text can no longer visually cover content cells further right.
  - Time values rounding up to the next midnight now roll the date forward (`1.9999999` + `m/d h:mm` renders `1/2 0:00`, matching Excel; previously `1/1 0:00`). The roll is derived by re-querying the engine with the rounded serial, so the 1900 leap-day ghost-date neighborhood stays correct too.
  - Column widths are honored when a `<col>` declares `width` without `customWidth="1"` (common in simplified third-party writers) — previously such columns fell back to the default 8.43 chars.
  - The visible-range math now subtracts the header offset instead of relying on the scroll buffer to mask a constant 24/40px window bias.

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
