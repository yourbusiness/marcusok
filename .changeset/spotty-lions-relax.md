---
"@marcusok/excel-exporter": patch
---

Fix concurrency and edge-case gaps found in a repo-wide review.

- Shared progress overlay (concurrency): a second export taking over an already-visible overlay now re-renders its texts immediately at `show` time (they used to stay on the previous caller until the new caller's first progress/phase event); while the overlay is still pending, the reveal deadline is now the latest `delayMs` among the concurrent callers (a caller with a longer delay used to be revealed early by the first caller's timer); a caller closing early now falls the driver back to the most recent still-active caller (the delayed reveal used to show an already-finished export's texts for the whole remaining export), and `minVisibleMs` is measured from the actual DOM reveal time (it used to be skipped entirely on the concurrent hand-over chain, letting the overlay flash).
- `toStr` now catches the `TypeError` thrown by `JSON.stringify` on circular-reference objects, honoring the documented "one bad object never loses a cell nor fails the export" behavior (previously a single circular value — e.g. mutually-referencing ORM entities — failed the whole export, twice on the worker-stream route which retries on the main thread).
- The injected index column no longer carries a default `width`, so stream exports with `indexColumn` enabled no longer print a misleading `features not supported (width)` warning when the user never configured any width. The visual default is unchanged (width 6 is applied where widths are consumed, on the Workbook path); an explicit `indexColumn.width` still warns on stream routes.
- `exportExcel(null)` now resolves with the structured `{ success: false }` error instead of rejecting with a raw TypeError, and an invalid `mode` value (e.g. `"Main"`) is rejected up front with a clear error instead of silently routing as `auto`.
