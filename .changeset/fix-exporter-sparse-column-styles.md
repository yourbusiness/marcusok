---
"@marcusok/excel-exporter": patch
---

Fix column styles silently landing on the wrong columns when earlier columns of a row are missing (sparse rows). Styles are now applied by each cell's real column reference instead of its position in the engine's densely-packed cell array, so a row missing column `b` no longer shifts `c`'s style onto `b`. The README also now notes that worker scripts are same-origin only (only the WASM binary may come from a plain CDN).
