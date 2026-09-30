# Package Relationships & Selection

The ecosystem publishes two packages across two categories (see [Ecosystem](/guide/) for the category definitions). This page answers the questions that come up when you pick packages: which one does what, what each of them bundles, whether the export and the preview can live on the same page, and how versions relate.

## The two published packages

| Package                                                 | Category | What it does                                                                        | Runtime dependencies                                     |
| ------------------------------------------------------- | -------- | ----------------------------------------------------------------------------------- | -------------------------------------------------------- |
| [`@marcusok/excel-exporter`](/packages/excel-exporter/) | Export   | Data → `.xlsx` download: styling, formatting, worker/stream routing, fallbacks      | `modern-xlsx` (types only; the engine itself is bundled) |
| [`@marcusok/excel-preview`](/packages/excel-preview/)   | Preview  | Read-only `.xlsx` / `.xlsm` / `.csv` rendering with worker parsing + virtual scroll | `modern-xlsx` (types only; the engine itself is bundled) |

Each package is self-contained at runtime. The two capability layers they are built on — the modern-xlsx engine layer (internally `xlsx-core`, WASM loader plus a stable re-export surface) and the progress overlay UI (internally `progress-overlay`) — are **private workspace packages**: they are bundled into the business packages' `dist` at build time and are not published to npm. There is nothing to install beyond the package you use, and no shared-package versions to reconcile.

Two consequences worth internalizing:

- **The exporter and the preview are siblings, not layers.** Neither depends on the other; you can install either alone. Nothing about the export engine leaks into the preview API, and the preview does not exist to serve the exporter.
- **Each package bundles the whole engine.** The modern-xlsx JS glue and the WASM binary ship inside each package's `dist` (and inside each package's worker — see below). The `modern-xlsx` entry that remains in `dependencies` is a **types-only** declaration: the published `.d.ts` files keep external type imports resolvable, but no modern-xlsx code is loaded at runtime.

### When to use which

- **`@marcusok/excel-exporter`** — you have data in the app (tables, arrays, ECharts options) and need a file the user downloads. Start at [Getting Started](/guide/01-getting-started); the convenience adapters are in [Table & ECharts](/packages/excel-exporter/api/05-table-and-echarts).
- **`@marcusok/excel-preview`** — you have a `.xlsx` the user uploaded and need to _show_ it, without editing. Start at [Quick Start](/packages/excel-preview/guide/01-quick-start).

## Using the exporter and the preview on the same page

Yes — they coexist without knowing about each other. Since each bundles its own copy of the engine, a page using **both** packages loads two engine instances on the main thread. What that means in practice:

- **Network cost is usually deduplicated.** Both packages ship the identical WASM binary; bundlers that name assets by content hash (Vite, webpack 5) emit one file for identical content, and the browser's HTTP cache covers the rest.
- **Memory is the real cost.** Each copy compiles and instantiates its own WASM module (a ~1.9 MB binary plus its linear memory). On a typical admin page — which uses the exporter _or_ the preview, rarely both — this never comes up; if you do run both on one performance-sensitive page, keep it in mind.
- **Workers were always like this.** A browser module worker cannot resolve bare specifiers (import maps do not apply inside `WorkerGlobalScope`), so each worker script has always been a single self-contained file bundling its own engine copy — `export.worker.js` (exporter) and `parse.worker.js` (preview) each carry one, unchanged from before.

### What is bundled, and what stays separate

| Thing                       | Per package or common?                                                                                                                                          |
| --------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Main-thread engine + WASM   | **Per package** — each package's `dist` carries its own engine copy and its own `modern-xlsx.wasm`                                                              |
| WASM loader & configuration | **Per package** — `configureWasm` / `getWasmLoader` from `@marcusok/excel-exporter` affect the exporter's copy; the preview's exports affect the preview's copy |
| Worker script               | **Per package** — `export.worker.js` (exporter) and `parse.worker.js` (preview) are different files, each shipped by its own package                            |
| Worker URL option           | **Per package field** — `workerUrl` is the export worker's, `parseWorkerUrl` is the parse worker's. One field could not serve both                              |
| Overlay UI                  | **Exporter only** — the preview does not bundle or use it                                                                                                       |

The loader row is the one that changes behaviour when both packages are present: **one `configureWasm` call configures one package's loader, not both.** If you self-host assets on a page using both packages, call `configureWasm` once per package you use, passing the same `wasmUrl`:

```ts
import { configureWasm as configureExportWasm } from "@marcusok/excel-exporter";
import { configureWasm as configurePreviewWasm } from "@marcusok/excel-preview";

configureExportWasm({ wasmUrl, workerUrl }); // exporter's loader + export worker
configurePreviewWasm({ wasmUrl, parseWorkerUrl }); // preview's loader + parse worker
```

Timeouts (`timeoutMs`, `maxRetries`, `workerTimeoutMs`) live on each loader and apply to that package's asset loads and worker dispatches. Fallbacks are per package as well: the exporter degrades to the style-less fast stream, the preview falls back to main-thread parsing; one package falling back does not change the other's route.

## The engine layer inside the packages

The engine integration — WASM loading (browser fetch / Node sync-init, timeout, retry, state machine), asset distribution and a stable re-export surface over modern-xlsx — lives in a single internal layer that both packages bundle. What that means for you:

- **You never think about engine versions.** There is no engine package to install or pin: upgrading the engine arrives as a release of the package you already use.
- **`configureWasm` / `getWasmLoader` are re-exported by both packages** with identical signatures and behaviour — import them from whichever package you already use; there is no second dependency to add for asset self-hosting.
- **The WASM binary ships under each package's `exports` map** (`@marcusok/excel-exporter/dist/modern-xlsx.wasm`, `@marcusok/excel-preview/dist/modern-xlsx.wasm`), so `?url` imports resolve with zero plugin configuration. Node locates the binary on disk next to the installed package and initializes synchronously — no boilerplate.

> Within one package, "single loader instance" is a module-instance guarantee: `getWasmLoader()` always returns the same object for that package. Across the two packages the loaders are separate by design (each bundles its own copy).

## The overlay inside the exporter

- Since **2.8.0**, every export shows the overlay by default. No extra import, no wrapper function.
- The `overlay` field is declared on `ExportOptions` and has three states — omitted/`true` (default export texts), `false` (fully off: nothing mounted, no extra frame yielded), or an options object (customize texts, `delayMs`, theme, …). All three, including the merge semantics for custom texts, are documented in [Progress Overlay](/packages/excel-exporter/guide/11-overlay#option-values) — this page does not repeat them.
- **The field exists on `ExportOptions` only.** `TableExportOptions` and `EChartsExportOptions` do not declare it, and their converters do not forward it, so the two convenience wrappers cannot customize or disable the overlay: it runs with its default (on) for them, and the main entry is the way to opt out. See [Table & ECharts → Relationship with exportExcel](/packages/excel-exporter/api/05-table-and-echarts#relationship-with-exportexcel).
- **Driving the overlay for non-Excel work.** The overlay is bundled into the exporter, so the public path is the `@marcusok/excel-exporter/overlay` subpath — it exports `showExportOverlay` (the generic overlay entry, Excel-agnostic) plus the handle types. The docs pattern is in [Overlay → Driving the overlay yourself](/packages/excel-exporter/guide/11-overlay#driving-the-overlay-yourself).

## Versions, releases and compatibility

- **Each package carries its own version, and the numbers are not aligned** — the two packages are released independently by Changesets (`.changeset/config.json` declares no `fixed` or `linked` groups). Do not read compatibility out of the version numbers: exporter 3.x and preview 2.x say nothing about each other.
- **They are both released from this one repository**, by the same CI, on the same `main`. Anything published close together was built from the same workspace state, which is the real compatibility statement.
- **There are no peer dependencies anywhere in the repo**, so there is no peer range for you to reconcile.
- **Engine upgrades are package releases.** Bumping modern-xlsx happens once in the internal engine layer, then ships as a new version of whichever business package consumes it — there is no separate engine release and no cross-package version coupling.

### How to judge whether two versions go together

Rather than a compatibility matrix, use the manifests — they are authoritative and always current:

1. Check what a package depends on: `npm view @marcusok/excel-exporter dependencies` (or look at the installed `package.json` under `node_modules`).
2. For behaviour changes between versions, read the package's `CHANGELOG.md` — it is generated by Changesets per package, so each entry maps to exactly the range you are upgrading.

## Which package should I install?

| I need to…                                                            | Install                                                                                                    |
| --------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| Export data to a downloadable `.xlsx`                                 | [`@marcusok/excel-exporter`](/packages/excel-exporter/)                                                    |
| Export from an Ant Design / Element Plus table or ECharts options     | `@marcusok/excel-exporter` — the adapters ship in the main entry                                           |
| Show an uploaded `.xlsx` / `.xlsm` / `.csv` read-only                 | [`@marcusok/excel-preview`](/packages/excel-preview/)                                                      |
| Parse a workbook headlessly (SSR, custom renderer, framework wrapper) | `@marcusok/excel-preview` — `parseWorkbookBytes`, no DOM required                                          |
| Do both — preview a file _and_ let the user export a corrected copy   | Both packages; each carries its own engine copy (see above)                                                |
| Drive the same progress overlay for non-Excel work                    | `@marcusok/excel-exporter` — the `/overlay` subpath exports the generic overlay entry                      |
| Configure WASM hosting (self-hosted / CDN copies)                     | Whichever package you use — its exported `configureWasm` covers its loader and worker URLs                 |
| Preview a `.csv` without touching the engine                          | `@marcusok/excel-preview` — its CSV path is pure JS; `.xlsx` / `.xlsm` parsing always uses the WASM engine |

## See also

- [Ecosystem](/guide/) — categories and roadmap.
- [Getting Started](/guide/01-getting-started) — the install line for the exporter and what it brings.
- [Assets & Self-hosting](/packages/excel-preview/guide/02-assets) — how WASM / worker assets resolve, and self-hosting.
- [Progress Overlay](/packages/excel-exporter/guide/11-overlay) — the `overlay` option and the legacy subpath.
