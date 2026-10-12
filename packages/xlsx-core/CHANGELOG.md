# @marcusok/xlsx-core

> **No longer published to npm** — since `@marcusok/excel-exporter` 3.0 / `@marcusok/excel-preview` 2.0 this package is a private build-time layer, bundled wholesale into both published packages' dist at build time. The 1.0.x versions below are the last published ones; subsequent changes ship inside the business packages and are recorded in their changelogs. The historical entries below describe the architecture as it was when this package was still published independently (e.g. "consumers keep it external" in 1.0.0) and no longer reflect the current bundling.

## 1.0.1

### Patch Changes

- 61fd7ba: Correct the stale entry-header comment (modern-xlsx is a declared runtime-free dependency kept for d.ts resolution, not a devDependency) and document that `workerTimeoutMs` must be positive — 0/negative values time out immediately rather than disabling the timeout.

## 1.0.0

### Major Changes

- be12efc: Initial release: the repo's single modern-xlsx engine layer, extracted from @marcusok/excel-exporter. Bundles the engine's runtime in at build time (nothing imports modern-xlsx at runtime, so the engines.node>=24 range is inert at runtime; the pinned `modern-xlsx` dependency declaration exists solely so TypeScript consumers can resolve the re-exported types in the published d.ts — note this does make the pinned version visible to installers, so engine-strict environments on Node < 24 will reject the install), owns WASM loading (browser fetch / Node sync-init dual path with timeout and retries, zero configuration in both; self-hosting options carry separate `workerUrl` / `parseWorkerUrl` fields so the exporter's and the preview's workers never cross-wire on a page using both), re-publishes the wasm binary under its own exports map, and exposes the shared API surface (readBuffer, formatCellRich, date/reference utilities, typed errors) plus the tsup esbuild plugins (`@marcusok/xlsx-core/tsup`) other packages reuse when bundling self-contained workers. Consumers keep it external in main builds, so pages using several @marcusok spreadsheet packages load one engine instance and one WASM binary.
