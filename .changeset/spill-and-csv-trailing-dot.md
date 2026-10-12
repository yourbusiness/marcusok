---
"@marcusok/excel-preview": patch
---

Align two edge behaviors with Excel semantics: text overflow is no longer blocked by value-less styled cells (an empty cell blocks spill in Excel regardless of formatting, so the spill limit now extends to the next cell with content or a merged region), and CSV values with a trailing decimal point (`"123."`) are now inferred as numbers, matching Excel's input-bar interpretation (leading-dot `".5"` and multi-dot `"1.2.3"` forms stay text).
