---
"@marcusok/excel-exporter": patch
---

Fix a batch of silent-data-loss and hang-prone edge cases found in a repo-wide review, and tighten a few validation/docs points.

- `exportEcharts` / `echartsToSheet`: multi-dimensional scatter data (`[x, y, ...dims]`, third dim onward typically driving symbolSize/visualMap) used to fall through to the name/value branch and silently produced `null` X/Y — coordinates are now taken from the first two dims, with a one-time console warning noting the dropped extra dims. Mixing 2-dim and multi-dim pairs no longer triggers the misleading "mixing scatter coordinate data with name/value data" error. Horizontal bar charts now take categories from `yAxis.data` (previously the category column was silently lost); multiple x/y axes are rejected with an explicit error instead of measuring series against `axis[0]`.
- `{ type: "number" }` columns now treat blank/whitespace-only strings as missing values (empty cell) on every path — `Number("") === 0` used to silently turn the most common database/form/CSV missing-value shape into a meaningful `0`. `{ type: "padding" }` likewise renders `null`/`undefined` as an empty cell instead of a padded fake code like `"00000"`.
- Cell values from objects whose `toJSON` returns `undefined` no longer disappear (and shift the following columns' styles by one cell): `toStr` falls back to a visible string.
- `exportExcelWithOverlay` no longer hangs forever when the page is hidden (background tab / minimized window): the pre-export double-RAF yield now races a short timer fallback, since browsers fully pause `requestAnimationFrame` on invisible pages.
- Pre-flight validation now covers style numeric fields (`alignment.textRotation` must be an integer 0–180 per its documented range, `font.size` a finite positive number) across all style locations (sheet-level, column-level incl. group columns, `indexColumn`) — invalid values used to fail deep inside the engine with an opaque serde error and degrade the whole export to the style-less stream.
- Build: the loader's Node auto-init `import("node:fs")` now survives esbuild's browser-platform specifier rewrite as a computed specifier, so published bundles no longer contain a bare `import('fs')` (the exact consumer-build warning shape the tsup config replaces for modern-xlsx).
