---
"@marcusok/excel-preview": patch
---

Fix a set of rendering and number-format issues found in a second review:

- Uppercase date/time format codes no longer misread `M` as a month: the minute/ month adjacency check was case-sensitive while format tokens are case-insensitive, so third-party codes like `HH:MM:SS` rendered as hour:month:second (e.g. `06:12:00` instead of `06:00:00`).
- Huge merged cells no longer break virtual scrolling. Merge recall used to widen the visible window to the whole merge rectangle, so a full-column merge (A1:A100000 — common in templates) made the renderer build every row header in the merge (100k+ DOM nodes). Only the merge's anchor cell is recalled now; headers stay at viewport size.
- Text spill is capped at the next non-empty cell in the row (Excel's truncation semantics) — overflowing text can no longer visually cover content cells further right.
- Time values rounding up to the next midnight now roll the date forward (`1.9999999` + `m/d h:mm` renders `1/2 0:00`, matching Excel; previously `1/1 0:00`). The roll is derived by re-querying the engine with the rounded serial, so the 1900 leap-day ghost-date neighborhood stays correct too.
- Column widths are honored when a `<col>` declares `width` without `customWidth="1"` (common in simplified third-party writers) — previously such columns fell back to the default 8.43 chars.
- The visible-range math now subtracts the header offset instead of relying on the scroll buffer to mask a constant 24/40px window bias.
