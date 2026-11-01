---
"@marcusok/excel-preview": patch
---

Decode XML entities in sheet names and font names (so `R&amp;D` displays and `setSheet("R&D")` works), keep numeric cells with the `@` text format unrounded, render the literal `-` in built-in accounting zero sections, map the engine's `SECURITY` (zip-bomb) error to `CORRUPT` with `domain`, and bound column-span layout work with a coverage budget so a crafted 100k-`<col>` file can no longer freeze the main thread.
