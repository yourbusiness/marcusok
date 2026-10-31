# Assets & Self-hosting

The preview ships two runtime assets: the WASM engine binary (1.9MB, gzip ≈ 650KB) and the self-contained parse worker. Both are located automatically in bundlers and Node; configuration is only needed for self-hosted copies.

## How assets resolve by default

| Asset        | Default location                                | Mechanism                                                                                             |
| ------------ | ----------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| WASM binary  | `@marcusok/excel-preview/dist/modern-xlsx.wasm` | `new URL(<file>, import.meta.url)` — bundlers emit a hashed asset; Node resolves through node_modules |
| Parse worker | `@marcusok/excel-preview/dist/parse.worker.js`  | Same pattern; the worker file is a single self-contained ESM (no sibling imports to 404)              |

The main thread needs no WASM at all in the browser, as long as the worker loads: parsing happens inside the worker (which initializes its own copy), and cell formatting — the only engine work done on the main thread — is pure JS. If the worker fails to construct/load, the session falls back to the main thread (and that route does load WASM — see [Quick Start](/packages/excel-preview/guide/01-quick-start)). The fallback is sticky for the session, but `terminateParseWorker()` clears it so the next parse can rebuild from the current configuration. On Node (no Worker) the main thread initializes the WASM synchronously from disk, zero boilerplate.

The engine is bundled into this package's `dist` at build time (before 2.0 it shipped in a separate `xlsx-core` package). A page using both `excel-exporter` and `excel-preview` carries one engine copy per package on the main thread — identical binaries, usually deduplicated as an asset by content-hash naming and the HTTP cache. Each package's self-contained worker bundles its own copy by necessity — browser module workers cannot resolve bare specifiers.

## Self-hosting / CDN

Point both URLs at your copies (the call shape matches the exporter's — each package configures its own loader):

```ts
import { configureWasm } from "@marcusok/excel-preview";

configureWasm({
  wasmUrl: "https://cdn.example.com/modern-xlsx.wasm",
  parseWorkerUrl: "https://cdn.example.com/parse.worker.js",
});
```

Note the field name: the parse worker uses its own `parseWorkerUrl` option — the loader's source is shared by the @marcusok packages, and each package's worker is a different script, so a single `workerUrl` field would cross-wire them (`workerUrl` is the _export_ worker's option, read by @marcusok/excel-exporter).

> **Worker scripts are same-origin only.** A browser rejects a cross-origin worker script at `Worker` construction (`SecurityError`) — a raw CDN URL for `parseWorkerUrl` will not load. The preview detects it, prints a console warning and falls back to main-thread parsing for the session, but to actually use the worker, serve it from your own origin (a local copy or a reverse proxy). Only the WASM binary can come from a plain CDN.

With bundler asset imports:

```ts
import wasmUrl from "@marcusok/excel-preview/dist/modern-xlsx.wasm?url";
import parseWorkerUrl from "@marcusok/excel-preview/dist/parse.worker.js?url";
configureWasm({ wasmUrl, parseWorkerUrl });
```

Notes carried over from the exporter's loader (they share it):

- Call `configureWasm` **before** the first preview; a WASM URL change after a successful load only takes effect in a fresh JS realm (page reload or a newly created worker).
- URL changes need the old worker gone: the parse worker reads `parseWorkerUrl` once at creation (the worker-side `initWasm` reads `wasmUrl` the same way). Call `terminateParseWorker()` (exported from the main entry) before the next parse — in-flight and queued parses are rejected without a main-thread rerun, the `workerBroken` permanent-fallback state is cleared, and the next parse creates a fresh worker from the current options. Recreating the preview is **not** enough: the worker is a module-level singleton shared across preview instances.
- Worker requests are serialized through a main-thread FIFO queue (one parse in flight). A request's timeout starts at dispatch, so queue wait is not counted; a timed-out parse terminates the worker and queued parses continue on a fresh instance. Aborting the active parse also terminates the worker; aborting a queued parse only removes that request.
- The worker falls back to the main-thread parse path when it fails to load — the preview keeps working, at the cost of blocking during large parses.
- Node runtime caveat: the bundled engine's `readBuffer` is ~10x slower on Node 22.x (and Node 24.14) than on Node 24.21 for numeric-heavy sheets. Functional support is Node >= 22 (CI covers it), but the performance numbers in this documentation assume the pinned development runtime (`.nvmrc`).
- Node bundler caveat: if a server build inlines this package and the WASM asset is not emitted next to the bundle, automatic disk lookup fails and parsing returns a `WASM` error (there is no silent style-less fallback for preview). Keep the package external, copy the asset, or point `configureWasm({ wasmUrl })` at it explicitly — Node accepts an HTTP(S) URL, a plain filesystem path, or a `file:` URL (the last two are read synchronously with `fs`).

## Loader defaults

The shared loader's tunables and their defaults (same for the exporter — one loader, one set of options):

| Option            | Default   | Meaning                                                                                                                                                                                                                                                    |
| ----------------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `timeoutMs`       | `10_000`  | Per-attempt WASM load timeout for the **main-thread** loader (fetch/instantiate). Worker parsing initializes inside the parse worker, where this value is not applied — that route is bounded by `workerTimeoutMs`.                                        |
| `maxRetries`      | `3`       | Total **main-thread** load attempts including the first (backoff 300ms / 600ms between retries); worker-side initialization is not retried.                                                                                                                |
| `workerTimeoutMs` | `120_000` | Parse-worker operation timeout. A timed-out parse terminates the shared worker and rejects its in-flight requests without a main-thread rerun; the next parse rebuilds a fresh worker. Must be > 0 — `0`/negative is "time out immediately", not "disable" |

## Vite dev-server caveat

Same as the exporter's: if the WASM request returns HTML (`content-type: text/html`, compile error mentions `expected magic word`), add the package to `optimizeDeps.exclude` so dependency pre-bundling stops intercepting the asset. See the [exporter's installation notes](/packages/excel-exporter/guide/02-installation#vite-dev-server-pre-bundling-caveat) for the full diagnosis.
