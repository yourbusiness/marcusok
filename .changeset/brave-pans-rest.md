---
"@marcusok/excel-exporter": major
---

Bundle the shared layers into the package and drop the same-scope runtime dependencies. `@marcusok/xlsx-core` (the modern-xlsx engine layer) and `@marcusok/progress-overlay` (the overlay UI) are no longer published to npm; both are private build-time layers now bundled into this package's `dist`, so installing `@marcusok/excel-exporter` brings the whole engine, overlay UI and WASM binary. Migration notes:

- **Remove any direct `@marcusok/xlsx-core` / `@marcusok/progress-overlay` imports** — both packages are no longer installable. `configureWasm` / `getWasmLoader` were already re-exported from the main entry; the overlay's standalone entry is now `showExportOverlay` on the `@marcusok/excel-exporter/overlay` subpath (also exports `nextPaint` and the overlay option/handle types under their legacy aliases).
- **WASM asset path**: the canonical self-hosting path is now `@marcusok/excel-exporter/dist/modern-xlsx.wasm?url` (previously `@marcusok/xlsx-core/dist/modern-xlsx.wasm`).
- **One loader per package**: `configureWasm` from this package configures this package's bundled loader only; `@marcusok/excel-preview` has its own. On a page using both, call it once per package (passing the same `wasmUrl` is fine — the binaries are identical).
- The only remaining `dependencies` entry is a types-only pinned `modern-xlsx` (published `.d.ts` files keep external type imports resolvable); nothing from it loads at runtime.
