---
"@marcusok/excel-preview": patch
---

Support OOXML's optional `row`/`cell` `r` attributes (derive positions from document order, accept `$A$1` references) instead of rejecting or dropping such workbooks; honor Excel's optional `#` decimal placeholders (`0.##`, `0.0#`, percentages) instead of keeping the engine's forced digits; recognize `[ColorN]` without a space; draw grid lines for empty cells in the visible window so sparse sheets no longer show broken cell borders; and refresh an existing overflowing cell's `max-width` while scrolling so long text is not clipped at a stale column boundary.
