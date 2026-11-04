---
"@marcusok/excel-preview": patch
---

Parse absolute references in `<mergeCell>` (`$A$1:$C$3`). The merge parser only accepted plain `A1:C3` refs and silently dropped absolute-form merges some producers write (the cell-reference path already accepted the same `$` form), losing the merged region entirely with no warning.
