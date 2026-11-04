---
"@marcusok/excel-preview": patch
---

Fix merged-cell rendering to match Excel:

- Text in a merged anchor cell never overflows. Horizontal merges were already clipped by the covered-cell blocker, but vertical merges (single column) spilled long text across the merge's column boundary into adjacent empty cells — merged cells clip at the merge boundary in Excel, in both directions.
- A merge whose whole region is empty (no value, no style — model sparsification drops the anchor cell) no longer renders as a grid hole: it now gets a full merge-sized placeholder with the outer grid frame, instead of a single stray one-cell box at the anchor.
