# Package Relationships & Selection

The ecosystem publishes five packages across two categories (see [Ecosystem](/guide/) for the category definitions). This page answers the questions that come up when you pick packages: which one does what, what each of them bundles, whether several can live on the same page, and how versions relate.

## The five published packages

| Package                                                 | Category | What it does                                                                                                 | Runtime dependencies                                  |
| ------------------------------------------------------- | -------- | ------------------------------------------------------------------------------------------------------------ | ----------------------------------------------------- |
| [`@marcusok/excel-exporter`](/packages/excel-exporter/) | Export   | Data → `.xlsx` download: styling, formatting, worker/stream routing, fallbacks                               | none (the engine itself is bundled)                   |
| [`@marcusok/excel-preview`](/packages/excel-preview/)   | Preview  | Read-only `.xlsx` / `.xlsm` / `.csv` rendering with worker parsing + virtual scroll                          | none (the engine itself is bundled)                   |
| [`@marcusok/word-preview`](/packages/word-preview/)     | Preview  | Read-only `.docx` / `.docm` rendering (docx-preview engine) with sniffing + hardened defaults                | none (docx-preview and jszip are bundled)             |
| [`@marcusok/pdf-preview`](/packages/pdf-preview/)       | Preview  | Read-only PDF rendering (pdf.js engine): lazy continuous scroll, selectable text layer                       | none (pdf.js is bundled; worker/cmaps ship in `dist`) |
| [`@marcusok/ppt-preview`](/packages/ppt-preview/)       | Preview  | Read-only PPTX rendering (Rust/WASM engine): lazy continuous scroll, 0-based navigation, password decryption | none (the engine + wasm are bundled)                  |

Each package is self-contained at runtime. The two Excel packages are built on two capability layers — the modern-xlsx engine layer (internally `xlsx-core`, WASM loader plus a stable re-export surface) and the progress overlay UI (internally `progress-overlay`) — which are **private workspace packages**: they are bundled into the Excel business packages' `dist` at build time (the engine layer into both Excel packages, the overlay into the exporter only — see the table below) and are not published to npm. The word, PDF and PPTX previews bundle their own engines the same way (docx-preview + jszip; pdf.js; a Rust/WASM OOXML engine) and have **zero** runtime dependencies. There is nothing to install beyond the package you use, and no shared-package versions to reconcile.

Three consequences worth internalizing:

- **The five packages are siblings, not layers.** None depends on another; you can install any of them alone. Nothing about the Excel engine leaks into the preview packages' APIs, and the previews do not exist to serve the exporter.
- **Each Excel package bundles the whole WASM engine.** The modern-xlsx JS glue and the WASM binary ship inside each Excel package's `dist` (and inside each package's worker — see below). Their published manifests have no modern-xlsx runtime dependency: the pinned entries in the repository are build-time dependencies used for engine/type resolution, and the published `.d.ts` files are self-contained. The word and PDF previews likewise have no runtime dependencies.
- **Same-page combinations are all fine** — see below.

### When to use which

- **`@marcusok/excel-exporter`** — you have data in the app (tables, arrays, ECharts options) and need a file the user downloads. Start at [Getting Started](/guide/01-getting-started); the convenience adapters are in [Table & ECharts](/packages/excel-exporter/api/05-table-and-echarts).
- **`@marcusok/excel-preview`** — you have a `.xlsx` the user uploaded and need to _show_ it, without editing. Start at [Quick Start](/packages/excel-preview/guide/01-quick-start).
- **`@marcusok/word-preview`** — you have a `.docx` the user uploaded and need to _show_ it, without editing. Start at [Quick Start](/packages/word-preview/guide/01-quick-start).
- **`@marcusok/pdf-preview`** — you have a `.pdf` the user uploaded and need to _show_ it, without editing. Start at [Quick Start](/packages/pdf-preview/guide/01-quick-start).
- **`@marcusok/ppt-preview`** — you have a `.pptx` the user uploaded and need to _show_ it, without editing. Start at [Quick Start](/packages/ppt-preview/guide/01-quick-start).

## Using several packages on the same page

Yes — they coexist without knowing about each other. The one combination with a real cost is the two Excel packages together: since each bundles its own copy of the engine, a page using **both** loads two engine instances on the main thread. What that means in practice:

- **Network cost is usually deduplicated.** Both Excel packages ship the identical WASM binary; bundlers that name assets by content hash (Vite, webpack 5) emit one file for identical content, and the browser's HTTP cache covers the rest.
- **Memory is the real cost.** Each copy compiles and instantiates its own WASM module (a ~1.9 MB binary plus its linear memory). On a typical admin page — which uses one package, rarely two — this never comes up; if you do run both on one performance-sensitive page, keep it in mind.
- **Workers were always like this.** A browser module worker cannot resolve bare specifiers (import maps do not apply inside `WorkerGlobalScope`), so each worker script has always been a single self-contained file bundling its own engine copy — `export.worker.js` (exporter) and `parse.worker.js` (preview) each carry one, unchanged from before. The PDF preview's worker (`pdf.worker.min.mjs`, upstream's own self-contained build) is independent of all of these.
- **The word preview adds none of this.** No WASM, no workers — it is plain JS/DOM and style-isolated per instance, so combining it with either or both Excel packages costs only its own bundle size. The PDF preview ships one pdf.js worker per instance (no WASM unless you opt into `wasmUrl`; cmaps/fonts are plain data files fetched over HTTP, never compiled), so it composes the same way. The PPTX preview ships one inlined parser worker + one ~2 MB WASM binary per instance (main-thread layout/paint), which composes like the Excel engine cost above — keep instance counts down on performance-sensitive pages.

### What is bundled, and what stays separate

| Thing                       | Per package or common?                                                                                                                                                                                                      |
| --------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Main-thread engine + WASM   | **Per Excel package** — each Excel package's `dist` carries its own engine copy and its own `modern-xlsx.wasm`; the word preview carries docx-preview + jszip instead (no WASM)                                             |
| WASM loader & configuration | **Per package** — `configureWasm` / `getWasmLoader` from `@marcusok/excel-exporter` affect the exporter's copy; the preview's exports affect the preview's copy                                                             |
| Worker script               | **Per package** — `export.worker.js` (exporter) and `parse.worker.js` (preview) are different files, each shipped by its own package; the word preview has no worker, the PDF preview ships upstream's `pdf.worker.min.mjs` |
| Worker URL option           | **Per package field** — `workerUrl` is the export worker's, `parseWorkerUrl` is the parse worker's. One field could not serve both                                                                                          |
| Overlay UI                  | **Exporter only** — the previews do not bundle or use it                                                                                                                                                                    |

The loader row is the one that changes behaviour when both Excel packages are present: **one `configureWasm` call configures one package's loader, not both.** If you self-host assets on a page using both Excel packages, call `configureWasm` once per Excel package you use, passing the same `wasmUrl`:

```ts
import { configureWasm as configureExportWasm } from "@marcusok/excel-exporter";
import { configureWasm as configurePreviewWasm } from "@marcusok/excel-preview";

configureExportWasm({ wasmUrl, workerUrl }); // exporter's loader + export worker
configurePreviewWasm({ wasmUrl, parseWorkerUrl }); // preview's loader + parse worker
```

Timeouts (`timeoutMs`, `maxRetries`, `workerTimeoutMs`) live on each loader and apply to that package's asset loads and worker dispatches. Fallbacks are per package as well: the exporter degrades to the style-less fast stream, the preview falls back to main-thread parsing; one package falling back does not change the other's route.

## The engine layer inside the Excel packages

The engine integration — WASM loading (browser fetch / Node sync-init, timeout, retry, state machine), asset distribution and a stable re-export surface over modern-xlsx — lives in a single internal layer that both Excel packages bundle. What that means for you:

- **You never think about engine versions.** There is no engine package to install or pin: upgrading the engine arrives as a release of the package you already use.
- **`configureWasm` / `getWasmLoader` are re-exported by both Excel packages** with identical signatures and behaviour — import them from whichever package you already use; there is no second dependency to add for asset self-hosting.
- **The WASM binary ships under each Excel package's `exports` map** (`@marcusok/excel-exporter/dist/modern-xlsx.wasm`, `@marcusok/excel-preview/dist/modern-xlsx.wasm`), so `?url` imports resolve with zero plugin configuration. Node locates the binary on disk next to the installed package and initializes synchronously — no boilerplate.

> Within one package, "single loader instance" is a module-instance guarantee: `getWasmLoader()` always returns the same object for that package. Across the two Excel packages the loaders are separate by design (each bundles its own copy).

## The overlay inside the exporter

- Since **2.8.0**, every export shows the overlay by default. No extra import, no wrapper function.
- The `overlay` field is declared on `ExportOptions` and has three states — omitted/`true` (default export texts), `false` (fully off: nothing mounted, no extra frame yielded), or an options object (customize texts, `delayMs`, theme, …). All three, including the merge semantics for custom texts, are documented in [Progress Overlay](/packages/excel-exporter/guide/11-overlay#option-values) — this page does not repeat them.
- **The field works on every entry.** Since 3.0.2, `TableExportOptions` and `EChartsExportOptions` declare `overlay` too, and both converters forward it to `exportExcel` — all three entries customize or disable the overlay identically. See [Table & ECharts → Relationship with exportExcel](/packages/excel-exporter/api/05-table-and-echarts#relationship-with-exportexcel).
- **Driving the overlay for non-Excel work.** The overlay is bundled into the exporter, so the public path is the `@marcusok/excel-exporter/overlay` subpath — it exports `showExportOverlay` (the generic overlay entry, Excel-agnostic) plus the handle types. The docs pattern is in [Overlay → Driving the overlay yourself](/packages/excel-exporter/guide/11-overlay#driving-the-overlay-yourself).

## Versions, releases and compatibility

- **Each package carries its own version, and the numbers are not aligned** — the five packages are released independently by Changesets (`.changeset/config.json` declares no `fixed` or `linked` groups). Do not read compatibility out of the version numbers: exporter 3.x, preview 2.x, word-preview 0.x and pdf-preview 0.x say nothing about each other.
- **They are all released from this one repository**, by the same CI, on the same `main`. Anything published close together was built from the same workspace state, which is the real compatibility statement.
- **There are no peer dependencies anywhere in the repo**, so there is no peer range for you to reconcile.
- **Engine upgrades are package releases.** A modern-xlsx bump lands in the internal engine layer and in the pinned build-time manifests of the two Excel business packages (three manifests move together; the runtime version is decided solely by the engine layer), then ships as a new version of whichever business package consumes it — there is no separate engine release and no cross-package version coupling. A docx-preview bump lands in the word preview alone; a pdfjs-dist bump lands in the PDF preview alone.

### How to judge whether two versions go together

Rather than a compatibility matrix, use the manifests — they are authoritative and always current:

1. Check what a package depends on: `npm view @marcusok/excel-exporter dependencies` (or look at the installed `package.json` under `node_modules`).
2. For behaviour changes between versions, read the package's `CHANGELOG.md` — it is generated by Changesets per package, so each entry maps to exactly the range you are upgrading.

## Which package should I install?

| I need to…                                                                     | Install                                                                                                               |
| ------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------- |
| Export data to a downloadable `.xlsx`                                          | [`@marcusok/excel-exporter`](/packages/excel-exporter/)                                                               |
| Export from an Ant Design / Element Plus table or ECharts options              | `@marcusok/excel-exporter` — the adapters ship in the main entry                                                      |
| Show an uploaded `.xlsx` / `.xlsm` / `.csv` read-only                          | [`@marcusok/excel-preview`](/packages/excel-preview/)                                                                 |
| Show an uploaded `.docx` / `.docm` read-only                                   | [`@marcusok/word-preview`](/packages/word-preview/)                                                                   |
| Show an uploaded `.pdf` read-only (lazy scroll, selectable text)               | [`@marcusok/pdf-preview`](/packages/pdf-preview/)                                                                     |
| Show an uploaded `.pptx` read-only (lazy scroll, slide navigation, decryption) | [`@marcusok/ppt-preview`](/packages/ppt-preview/)                                                                     |
| Parse a workbook headlessly (SSR, custom renderer, framework wrapper)          | `@marcusok/excel-preview` — `parseWorkbookBytes`, no DOM required                                                     |
| Preview a Word document with visual fidelity (pagination, footnotes)           | `@marcusok/word-preview` — note it is main-thread rendering, see its [Limits](/packages/word-preview/guide/02-limits) |
| Do both — preview a file _and_ let the user export a corrected copy            | The two Excel packages; each carries its own engine copy (see above)                                                  |
| Drive the same progress overlay for non-Excel work                             | `@marcusok/excel-exporter` — the `/overlay` subpath exports the generic overlay entry                                 |
| Configure WASM hosting (self-hosted / CDN copies)                              | Whichever Excel package you use — its exported `configureWasm` covers its loader and worker URLs                      |
| Preview a `.csv` without touching the engine                                   | `@marcusok/excel-preview` — its CSV path is pure JS; `.xlsx` / `.xlsm` parsing always uses the WASM engine            |

## See also

- [Ecosystem](/guide/) — categories and roadmap.
- [Getting Started](/guide/01-getting-started) — the install line for the exporter and what it brings.
- [Assets & Self-hosting](/packages/excel-preview/guide/02-assets) — how WASM / worker assets resolve, and self-hosting.
- [Progress Overlay](/packages/excel-exporter/guide/11-overlay) — the `overlay` option and the legacy subpath.
