# Assets & Self-hosting

The preview ships two runtime assets: the WASM engine binary (1.9MB, gzip ≈ 650KB) and the self-contained parse worker. Both are located automatically in bundlers and Node; configuration is only needed for self-hosted copies.

## How assets resolve by default

| Asset        | Default location                               | Mechanism                                                                                             |
| ------------ | ---------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| WASM binary  | `@marcusok/xlsx-core/dist/modern-xlsx.wasm`    | `new URL(<file>, import.meta.url)` — bundlers emit a hashed asset; Node resolves through node_modules |
| Parse worker | `@marcusok/excel-preview/dist/parse.worker.js` | Same pattern; the worker file is a single self-contained ESM (no sibling imports to 404)              |

The main thread needs no WASM at all in the browser: parsing happens inside the worker (which initializes its own copy), and cell formatting — the only engine work done on the main thread — is pure JS. On Node (no Worker) the main thread initializes the WASM synchronously from disk, zero boilerplate.

The engine lives in the shared [@marcusok/xlsx-core](https://www.npmjs.com/package/@marcusok/xlsx-core) package: a page using both `excel-exporter` and `excel-preview` loads one engine instance and one WASM binary on the main thread. Each package's self-contained worker bundles its own copy by necessity — browser module workers cannot resolve bare specifiers.

## Self-hosting / CDN

Point both URLs at your copies (same call as the exporter's — the loader is shared):

```ts
import { configureWasm } from "@marcusok/excel-preview";

configureWasm({
  wasmUrl: "https://cdn.example.com/modern-xlsx.wasm",
  parseWorkerUrl: "https://cdn.example.com/parse.worker.js",
});
```

Note the field name: the parse worker uses its own `parseWorkerUrl` option — the shared loader serves several @marcusok packages, and each package's worker is a different script, so a single `workerUrl` field would cross-wire them (`workerUrl` is the _export_ worker's option, read by @marcusok/excel-exporter).

> **Worker scripts are same-origin only.** A browser rejects a cross-origin worker script at `Worker` construction (`SecurityError`) — a raw CDN URL for `parseWorkerUrl` will not load. The preview detects it, prints a console warning and falls back to main-thread parsing for the session, but to actually use the worker, serve it from your own origin (a local copy or a reverse proxy). Only the WASM binary can come from a plain CDN.

With bundler asset imports:

```ts
import wasmUrl from "@marcusok/xlsx-core/dist/modern-xlsx.wasm?url";
import parseWorkerUrl from "@marcusok/excel-preview/dist/parse.worker.js?url";
configureWasm({ wasmUrl, parseWorkerUrl });
```

Notes carried over from the exporter's loader (they share it):

- Call `configureWasm` **before** the first preview; a WASM URL change after a successful load only takes effect in a fresh JS realm (page reload or a newly created worker).
- The worker falls back to the main-thread parse path when it fails to load — the preview keeps working, at the cost of blocking during large parses.

## Loader defaults

The shared loader's tunables and their defaults (same for the exporter — one loader, one set of options):

| Option            | Default   | Meaning                                                                                                                                                                                                                                                    |
| ----------------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `timeoutMs`       | `10_000`  | Per-attempt WASM load timeout (fetch/instantiate)                                                                                                                                                                                                          |
| `maxRetries`      | `3`       | Total load attempts including the first (backoff 300ms / 600ms between retries)                                                                                                                                                                            |
| `workerTimeoutMs` | `120_000` | Parse-worker operation timeout. A timed-out parse terminates the shared worker and rejects its in-flight requests without a main-thread rerun; the next parse rebuilds a fresh worker. Must be > 0 — `0`/negative is "time out immediately", not "disable" |

## Vite dev-server caveat

Same as the exporter's: if the WASM request returns HTML (`content-type: text/html`, compile error mentions `expected magic word`), add the package to `optimizeDeps.exclude` so dependency pre-bundling stops intercepting the asset. See the [exporter's installation notes](/packages/excel-exporter/guide/02-installation#vite-dev-server-pre-bundling-caveat) for the full diagnosis.
