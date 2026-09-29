---
"@marcusok/excel-exporter": minor
---

Internal refactor: the WASM loader, engine bundling and shared re-export surface moved into the new `@marcusok/xlsx-core` package, which is now this package's single runtime dependency. The public API is unchanged (`configureWasm` / `getWasmLoader` / `LoaderOptions` / `LoadState` are re-exported from the shared core with identical signatures and behavior), worker threading and all export modes are unchanged, and pages combining @marcusok packages now share one engine instance and one WASM binary on the main thread. The `@marcusok/excel-exporter/dist/modern-xlsx.wasm` asset path still resolves (deprecated) — prefer `@marcusok/xlsx-core/dist/modern-xlsx.wasm`.
