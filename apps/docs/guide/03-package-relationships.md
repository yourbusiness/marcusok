# Package Relationships & Selection

The ecosystem publishes four packages across three categories (see [Ecosystem](/guide/) for the category definitions). This page answers the questions that come up when you pick packages: which one does what, what pulls in what, whether the export and the preview can live on the same page, and how versions relate.

## The four published packages

| Package                                                     | Category | What it does                                                                                           | Runtime dependencies                                  |
| ----------------------------------------------------------- | -------- | ------------------------------------------------------------------------------------------------------ | ----------------------------------------------------- |
| [`@marcusok/excel-exporter`](/packages/excel-exporter/)     | Export   | Data → `.xlsx` download: styling, formatting, worker/stream routing, fallbacks                         | `@marcusok/xlsx-core`, `@marcusok/progress-overlay`   |
| [`@marcusok/excel-preview`](/packages/excel-preview/)       | Preview  | Read-only `.xlsx` / `.xlsm` / `.csv` rendering with worker parsing + virtual scroll                    | `@marcusok/xlsx-core`                                 |
| [`@marcusok/xlsx-core`](/packages/xlsx-core/)               | Shared   | The repo's single modern-xlsx engine layer: WASM loading, asset distribution, stable re-export surface | `modern-xlsx` (bundled into its `dist` at build time) |
| [`@marcusok/progress-overlay`](/packages/progress-overlay/) | Shared   | Framework-agnostic full-screen progress overlay (spinner / percentage bar, themes)                     | none                                                  |

The dependency graph is small and deliberate:

```
@marcusok/excel-exporter ──┬──> @marcusok/xlsx-core
                            └──> @marcusok/progress-overlay

@marcusok/excel-preview  ───────> @marcusok/xlsx-core
```

Three consequences worth internalizing:

- **The exporter and the preview are siblings, not layers.** Neither depends on the other; you can install either alone. Nothing about the export engine leaks into the preview API, and the preview does not exist to serve the exporter.
- **`@marcusok/xlsx-core` is the only shared engine.** Both app-facing packages depend on it directly — that is what makes their engine identical (see _Sharing one engine_ below).
- **`@marcusok/progress-overlay` is a leaf.** It has no dependencies at all, knows nothing about Excel, and is used by the exporter only — the preview does not pull it in. It is a general-purpose UI package that happens to be what the exporter's `overlay` option drives ([Progress Overlay](/packages/excel-exporter/guide/11-overlay)).

### When to use which

- **`@marcusok/excel-exporter`** — you have data in the app (tables, arrays, ECharts options) and need a file the user downloads. Start at [Getting Started](/guide/01-getting-started); the convenience adapters are in [Table & ECharts](/packages/excel-exporter/api/05-table-and-echarts).
- **`@marcusok/excel-preview`** — you have a `.xlsx` the user uploaded and need to _show_ it, without editing. Start at [Quick Start](/packages/excel-preview/guide/01-quick-start).
- **`@marcusok/xlsx-core`** — normally installed _for_ you as a dependency. Reach for it directly only when you are building your own integration on the modern-xlsx surface, or when you want to configure the WASM loader without importing a document package — the loader lives here (see [configureWasm](/packages/excel-exporter/api/01-export-excel#configurewasm)).
- **`@marcusok/progress-overlay`** — you want the same overlay for work that has nothing to do with Excel. Usage is in the [package guide](/packages/progress-overlay/guide/01-usage).

## Using the exporter and the preview on the same page

Yes — they are designed to coexist, and the two packages are aware of each other only through the shared engine. If you install both, you get **one engine instance and one WASM binary on the main thread**, not two: the main-thread builds of both packages keep `@marcusok/xlsx-core` external (rather than bundling their own copy), so both resolve to the same module at runtime. This is stated on the preview side in [Install](/packages/excel-preview/#install) and explained in [Assets & Self-hosting](/packages/excel-preview/guide/02-assets#how-assets-resolve-by-default).

### What is shared, and what is not

| Thing                            | Shared or per-package?                                                                                                               |
| -------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| Main-thread engine + WASM        | **Shared** — one copy of `@marcusok/xlsx-core`, one WASM binary on the main thread                                                   |
| WASM loader & configuration      | **Shared** — `configureWasm` / `getWasmLoader` are the same objects in both packages                                                 |
| Worker script                    | **Per package** — `export.worker.js` (exporter) and `parse.worker.js` (preview) are different files, each shipped by its own package |
| Worker URL option                | **Per package field** — `workerUrl` is the export worker's, `parseWorkerUrl` is the parse worker's. One field could not serve both   |
| Engine copy _inside_ each worker | **Per package** — unavoidable, see below                                                                                             |
| Overlay UI                       | **Exporter only** — the preview has no overlay dependency                                                                            |

The worker detail is the part that surprises people. A browser module worker cannot resolve bare specifiers (import maps do not apply inside `WorkerGlobalScope`), so each worker script must be a single self-contained file — which means each one **bundles its own copy of the engine**. The duplication is inherent to the worker boundary, not an oversight; the shared-engine guarantee is a main-thread guarantee. The preview's assets page spells this out, and the worker construction code refuses to share a single URL field for exactly this reason ([worker-client.ts](https://github.com/yourbusiness/marcusok/blob/main/packages/excel-preview/src/worker/worker-client.ts) reads `parseWorkerUrl`, `worker-exporter.ts` reads `workerUrl`).

So on a page using both packages you can see, at most: one main-thread engine, one export worker, one parse worker. In practice a single page rarely runs both worker paths at once.

### Practical consequences

- **One `configureWasm` call configures both.** There is a single loader, so a self-hosting setup passes `wasmUrl` plus both worker URLs in one call — see the preview's [Self-hosting / CDN](/packages/excel-preview/guide/02-assets) section, which is written for exactly this situation.
- **Timeouts are shared settings.** `timeoutMs`, `maxRetries` and `workerTimeoutMs` live on that one loader, so they apply to whichever package loads assets or dispatches a worker.
- **Fallbacks are per package.** Each package degrades on its own terms: the exporter has the style-less fast stream, the preview falls back to main-thread parsing. One package falling back does not change the other's route.
- **You do not need the exporter to parse, or the preview to write.** Sharing an engine does not mean sharing an API.

## Sharing one engine (`xlsx-core`)

The engine package has its own [documentation page](/packages/xlsx-core/), including a [relationship summary](/packages/xlsx-core/#relationship-with-the-two-business-packages) for the same three points covered here. What follows is the version/release angle, which matters when you are choosing packages rather than building on the engine.

### It is a regular dependency, not a peer dependency

Both business packages declare the engine as a plain `dependencies` entry:

- `packages/excel-exporter/package.json` — `"@marcusok/xlsx-core": "workspace:*"` (alongside `@marcusok/progress-overlay`)
- `packages/excel-preview/package.json` — `"@marcusok/xlsx-core": "workspace:*"` (its only runtime dependency)

Nothing in the repo declares `peerDependencies`. `workspace:*` is a development-time protocol; it is rewritten to a concrete version when the package is published (today both packages pin `@marcusok/xlsx-core` to an exact version in their published manifests — check yours with `npm view @marcusok/excel-exporter dependencies`).

That choice has a direct consequence for consumers: **you never have to think about which engine version to install.** `pnpm add @marcusok/excel-preview` brings the engine it was built and tested against; upgrading the engine is a matter of upgrading the package you already use, not of adding a third entry to your own dependency list, and it can never drift into a range conflict with the exporter's engine.

### Why the re-export surface gives a single version source

The modern-xlsx glue and `fflate` are bundled _into_ `xlsx-core`'s `dist` at build time, and `xlsx-core` is nothing but a loader plus a re-export surface over modern-xlsx. The business packages import the engine only through that surface — the exporter even re-exports `configureWasm` / `getWasmLoader` from it so that its public API is unchanged. Consequently:

- There is exactly one place in the repo that names a modern-xlsx version.
- Both packages see the same engine build, the same error types and the same WASM binary URL default.

### `getWasmLoader()` is literally the same object

`xlsx-core` exposes one module-level loader instance and hands it out from `getWasmLoader()`. Both packages re-export that function, so importing `configureWasm` or `getWasmLoader` from `@marcusok/excel-exporter` and from `@marcusok/excel-preview` configures and observes **the same loader** — as the preview's API reference notes ([createPreview → Other exports](/packages/excel-preview/api/01-create-preview#other-exports)): "importing them from either package configures the same shared loader."

> A single instance is a module-instance guarantee: it holds as long as your bundler resolves one copy of `@marcusok/xlsx-core`. The exact-version pin above is what makes that the normal case; if you find yourself with two copies of the package in one bundle, you have two engines and two loader states.

### The build-time half of `xlsx-core`

`xlsx-core` is also a _build-time_ dependency inside this repo: it publishes a `@marcusok/xlsx-core/tsup` subpath with the shared esbuild plugins both business packages use in their `tsup.config.ts`, so the two packages cannot drift in how they bundle the engine (see the [exports map](/packages/xlsx-core/#exports-map)). Nothing in your application imports that subpath.

Each business package also re-forwards the WASM binary into its own `dist` (a post-build copy step). That is a compatibility forward for older deep-import paths; the canonical location is the core package's `dist/modern-xlsx.wasm`, which is what the loader defaults to.

## `progress-overlay` and the exporter

The overlay is the one place where a shared package is wired _into_ a business package's default behaviour.

- Since **2.8.0**, every export shows the overlay by default. No extra import, no wrapper function.
- The `overlay` field is declared on `ExportOptions` and has three states — omitted/`true` (default export texts), `false` (fully off: nothing mounted, no extra frame yielded), or a `ProgressOverlayOptions` object (customize texts, `delayMs`, theme, …). All three, including the merge semantics for custom texts, are documented in [Progress Overlay](/packages/excel-exporter/guide/11-overlay#option-values) — this page does not repeat them.
- **The field exists on `ExportOptions` only.** `TableExportOptions` and `EChartsExportOptions` do not declare it, and their converters do not forward it, so the two convenience wrappers cannot customize or disable the overlay: it runs with its default (on) for them, and the main entry is the way to opt out. See [Table & ECharts → Relationship with exportExcel](/packages/excel-exporter/api/05-table-and-echarts#relationship-with-exportexcel).
- **The pre-2.8 subpath still exists for compatibility.** `@marcusok/excel-exporter/overlay` now holds a thin wrapper (`exportExcelWithOverlay`) equivalent to `exportExcel({ ...options, overlay })`; note that the text shape and handle method names changed when the feature moved into the shared package. Migration notes are in [Progress Overlay → Legacy subpath](/packages/excel-exporter/guide/11-overlay#legacy-subpath).
- **The package stands alone.** If your own flow needs the same overlay, import `@marcusok/progress-overlay` and drive the handle directly — the docs pattern is in [Overlay → Driving the overlay yourself](/packages/excel-exporter/guide/11-overlay#driving-the-overlay-yourself), and the package has no Excel dependency whatsoever.

## Versions, releases and compatibility

- **Each package carries its own version, and the numbers are not aligned** — the four packages are released independently by Changesets (`.changeset/config.json` declares no `fixed` or `linked` groups). Do not read compatibility out of the version numbers: exporter 2.x, preview 1.x, core 1.x and overlay 1.x say nothing about each other.
- **They are all released from this one repository**, by the same CI, on the same `main`. Anything published close together was built from the same workspace state, which is the real compatibility statement — the packages are not independently-maintained projects that happen to share a scope.
- **There are no peer dependencies anywhere in the repo**, so there is no peer range for you to reconcile. A business package pins the engine version it ships with (see above).
- **Engine upgrades propagate upward.** When `xlsx-core` is released, the dependent packages' dependency entry is moved with it and they are released in turn — the config sets `updateInternalDependencies: "patch"` (`.changeset/config.json`), so an engine release reaches consumers as a new exporter/preview version rather than as a change they have to make themselves. Internal dependencies also drive build order in this repo (`^build` in Turborepo), so a business package is never built against a stale core.

### How to judge whether two versions go together

Rather than a compatibility matrix, use the manifests — they are authoritative and always current:

1. Check what a package pins: `npm view @marcusok/excel-exporter dependencies` (or look at the installed `package.json` under `node_modules`).
2. Check that the engine version it pins is the one installed — with a plain `pnpm add` of one of the document packages there is nothing to reconcile; only a hand-pinned engine in your own manifest can conflict.
3. For behaviour changes between versions, read the package's `CHANGELOG.md` — it is generated by Changesets per package, so each entry maps to exactly the range you are upgrading.

## Which package should I install?

| I need to…                                                            | Install                                                                                                    |
| --------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| Export data to a downloadable `.xlsx`                                 | [`@marcusok/excel-exporter`](/packages/excel-exporter/)                                                    |
| Export from an Ant Design / Element Plus table or ECharts options     | `@marcusok/excel-exporter` — the adapters ship in the main entry                                           |
| Show an uploaded `.xlsx` / `.xlsm` / `.csv` read-only                 | [`@marcusok/excel-preview`](/packages/excel-preview/)                                                      |
| Parse a workbook headlessly (SSR, custom renderer, framework wrapper) | `@marcusok/excel-preview` — `parseWorkbookBytes`, no DOM required                                          |
| Do both — preview a file _and_ let the user export a corrected copy   | Both document packages; the engine and its WASM binary are loaded once                                     |
| Drive the same progress overlay for non-Excel work                    | [`@marcusok/progress-overlay`](/packages/progress-overlay/)                                                |
| Configure WASM hosting centrally for whichever packages you use       | Either document package (the exported `configureWasm` is shared), or `@marcusok/xlsx-core` itself          |
| Build on top of modern-xlsx directly (own builder, own read path)     | [`@marcusok/xlsx-core`](/packages/xlsx-core/)                                                              |
| Preview a `.csv` without touching the engine                          | `@marcusok/excel-preview` — its CSV path is pure JS; `.xlsx` / `.xlsm` parsing always uses the WASM engine |

## See also

- [Ecosystem](/guide/) — categories and roadmap.
- [Getting Started](/guide/01-getting-started) — the install line for the exporter and what it brings.
- [@marcusok/xlsx-core](/packages/xlsx-core/) — the engine layer, its loader and its exports map.
- [Assets & Self-hosting](/packages/excel-preview/guide/02-assets) — the loader both packages share, and how assets resolve.
- [Progress Overlay](/packages/excel-exporter/guide/11-overlay) — the `overlay` option and the legacy subpath.
