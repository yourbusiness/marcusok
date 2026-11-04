---
"@marcusok/excel-preview": patch
---

Rebuild an already-rendered cell into its text-overflow form once the overflow cap grows past the cell's own width. A long-text cell entering the viewport right next to a hidden (zero-width) column was created without the overflow span — its cap equaled its own width at entry — and as the window scrolled past the hidden column the cap grew but the existing element kept the plain clipped form, so the text stayed cut off at the column edge for as long as the cell remained visible (Excel keeps overflowing across hidden empty columns). Style-only blank cells (null value with a style index) are excluded from the conversion so they are not rebuilt every scroll frame.
