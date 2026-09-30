# @marcusok/xlsx-core

The repo's single integration point around [modern-xlsx](https://github.com/ABCrimson/modern-xlsx) (Rust + WASM): WASM loading, asset distribution and a stable re-export surface of the engine. It adds no behavior of its own beyond the loader — everything else it exposes is the engine's public API, re-published so the @marcusok packages never depend on `modern-xlsx` directly.

The engine's runtime is bundled **into** this package's `dist/` at build time, so consumers see zero external runtime dependencies and are immune to modern-xlsx's own `engines.node >= 24` declaration. The pinned `modern-xlsx` dependency in `package.json` exists only so TypeScript consumers can resolve the re-exported types (see [Package integration](/packages/xlsx-core/guide/03-package-integration)).

## What is inside

| Layer              | What it provides                                                                                                                                                                                 |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| WASM loader        | `configureWasm()` / `getWasmLoader()` / `WasmLoader` / `defaultWasmUrl()` — browser fetch and Node sync-init dual path, per-attempt timeout and retries, one shared loader instance              |
| Asset distribution | `dist/modern-xlsx.wasm` (~1.9MB), re-published under the exports map so `@marcusok/xlsx-core/dist/modern-xlsx.wasm?url` resolves in Vite / webpack 5                                             |
| Engine re-export   | The modern-xlsx runtime surface the @marcusok packages share: `Workbook`, `readBuffer`, reference/format/date utilities, typed errors ([full list](/packages/xlsx-core/guide/02-engine-surface)) |
| Build plugins      | `@marcusok/xlsx-core/tsup` — the two esbuild plugins downstream packages reuse when bundling ([details](/packages/xlsx-core/guide/03-package-integration))                                       |

## When to use it directly

- **You are building your own renderer, writer or data pipeline.** The exporter and the preview model data their own way; if you want the engine plus its primitives — `readBuffer` into a `Workbook`, `formatCellRich`, `decodeRange`, `serialToDate` — this is the layer to build on.
- **You need the typed error surface.** `ModernXlsxError` carries a machine-readable `code`; the four constants this package re-exports cover the format/password/init cases the business packages branch on (see [Engine surface](/packages/xlsx-core/guide/02-engine-surface#typed-errors)).
- **You target something other than Excel.** The engine reads and writes OOXML; anything that only needs the reader (a diff tool, a validator, an import pipeline) can use it without pulling in the export or preview UI.
- **You want one engine instance per page.** Importing the loader from here (or configuring it here) is the same singleton the business packages use, so a page that also uses them still loads a single WASM module.

## When not to

Exporting a workbook? Use [@marcusok/excel-exporter](/packages/excel-exporter/). Previsualizing one? Use [@marcusok/excel-preview](/packages/excel-preview/). Both **re-export** `configureWasm` / `getWasmLoader` (they are `@marcusok/xlsx-core`'s own functions, re-published), so importing them from either package configures the same shared loader — there is no reason to add this dependency just for asset self-hosting.

## Install

```bash
pnpm add @marcusok/xlsx-core
```

Requirements: Node `>= 22` (this package's `engines`), browsers with WebAssembly support. No runtime dependency is pulled in — the engine ships inside this package's `dist/`.

## Exports map

| Subpath                                     | Resolves to                         | Use it for                                                                                 |
| ------------------------------------------- | ----------------------------------- | ------------------------------------------------------------------------------------------ |
| `@marcusok/xlsx-core`                       | `dist/index.js` + `dist/index.d.ts` | The loader and the engine re-export surface                                                |
| `@marcusok/xlsx-core/tsup`                  | `dist/tsup.js` + `dist/tsup.d.ts`   | `rewriteWasmBgUrl` / `dropNodeFsPromises`, for a downstream package's own `tsup.config.ts` |
| `@marcusok/xlsx-core/dist/modern-xlsx.wasm` | the WASM binary                     | `?url` imports in bundlers that support asset imports (`...modern-xlsx.wasm?url`)          |
| `@marcusok/xlsx-core/package.json`          | the manifest                        | Tooling that reads the version / engines range instead of importing the module             |

## Relationship with the two business packages

- **One loader.** `getWasmLoader()` returns the same `WasmLoader` object for every importer: configuring the WASM URL through `@marcusok/excel-exporter` or `@marcusok/excel-preview` configures the one this package owns.
- **One WASM binary on the main thread.** Both business packages keep `@marcusok/xlsx-core` external in their main builds, so a page using several @marcusok packages compiles and instantiates the engine once.
- **Workers are the exception.** Each package's self-contained worker bundles its own engine copy — browser module workers cannot resolve bare specifiers, so the duplication there is inherent (the exporter's and the preview's worker paths are described in their own guides).

Configuration is optional in both environments: assets resolve to their shipped locations by default. Start with the [loader guide](/packages/xlsx-core/guide/01-loader); the [API reference](/packages/xlsx-core/api/01-api) is the terse version of the same surface.
