---
"@marcusok/excel-preview": patch
---

Cap the per-merge cell expansion in layout building: a pathological whole-sheet merge (`A1:XFD1048576`, ~1.7e10 cells) previously froze the main thread while populating the covered-cell index. Merges over 1M cells now keep only their anchor entry (the anchor still renders at full span); normal large merges such as full-column merges are unaffected.
