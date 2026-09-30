---
"@marcusok/excel-preview": major
---

Bundle the engine layer into the package and drop the same-scope runtime dependency. `@marcusok/xlsx-core` is no longer published to npm; it is a private build-time layer now bundled into this package's `dist`, so installing `@marcusok/excel-preview` brings the whole parsing engine and the WASM binary. Migration notes:

- **Remove any direct `@marcusok/xlsx-core` imports** — the package is no longer installable. `configureWasm` / `getWasmLoader` were already re-exported from the main entry.
- **WASM asset path**: the canonical self-hosting path is now `@marcusok/excel-preview/dist/modern-xlsx.wasm?url` (previously `@marcusok/xlsx-core/dist/modern-xlsx.wasm`).
- **One loader per package**: `configureWasm` from this package configures this package's bundled loader only; `@marcusok/excel-exporter` has its own. On a page using both, call it once per package (passing the same `wasmUrl` is fine — the binaries are identical).
- The only remaining `dependencies` entry is a types-only pinned `modern-xlsx` (published `.d.ts` files keep external type imports resolvable); nothing from it loads at runtime.
