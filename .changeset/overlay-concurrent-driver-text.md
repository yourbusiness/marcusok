---
"@marcusok/excel-exporter": patch
---

Fix the shared progress overlay rendering the wrong caller's text during concurrent exports, and clarify a few docs/config points found in a repo-wide review.

The overlay is a single shared DOM node. Its `title` and `hint` were written once by `createDom` (the first caller's texts) and never refreshed, and the delayed-reveal timer carried its initiating caller's `text` closure — so when export A armed the timer, closed, and export B took over as driver, the overlay could appear with A's title while showing B's state. The state machine and its texts are now paired as one `OverlayDriver`, `reveal()` always renders the current driver's texts, and `render()` refreshes `title`/`hint` on change alongside the label.

Also:

- `WorkbookBuilder.addSheet` now validates the sheet name before building the row array, so an invalid name no longer pays for the full O(rows×cols) mapping first (no behavior change — same error, same message).
- JSDoc on `SheetConfig.headerStyle` / `ColumnConfig.headerStyle` now states explicitly that a column-level `headerStyle` replaces the sheet-level one wholesale (unlike `style`, which deep-merges over `dataStyle`).
- Docs: the API page listed "two" code-splitting subpaths — now three, including `@marcusok/excel-exporter/overlay`; the installation/fallback guides note the loader's late-success self-heal exception to the "WASM load previously failed" behaviour; the overlay guide's concurrency section documents text ownership and the shared `delayMs` gate (en + zh).
