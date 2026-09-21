# Installation & Configuration

## Requirements

- Node `>= 22` (the package's `engines` requirement); any package manager works — examples use pnpm, npm / yarn are equivalent
- Browsers need WebAssembly support (all modern browsers)

## Install

```bash
pnpm add @marcusok/excel-exporter
```

That is the entire setup. The package has **zero runtime dependencies**: the export engine (modern-xlsx JS glue + fflate) is bundled in at build time, and the WASM binary ships under this package's own `exports` map — there is no engine package to install and no optional fallback package. Bundler config is needed in exactly one case: Vite's dev server (see the [pre-bundling caveat](#vite-dev-server-pre-bundling-caveat) below).

## Browser: assets resolve automatically

Two files ship alongside the code and are located by default, so a plain `import { exportExcel } from "@marcusok/excel-exporter"` works out of the box:

| Asset              | Description                                                                                                                                                                 |
| ------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `modern-xlsx.wasm` | WASM core (~1.9MB); needed by the styled routes (main / worker + Workbook) — the Fast stream routes (explicit `mode: "stream"`, or auto/Node ≥ 50,000 rows) do not use WASM |
| `export.worker.js` | Self-contained worker entry (a single ESM file, zero imports); needed only when exports enter a Worker (auto ≥ 20,000 rows, or explicit worker/stream mode)                 |

Resolution order:

1. **Bundlers** — both default to `new URL(<file>, import.meta.url)` next to the package entry. Production builds rewrite the expression and emit the file as a hashed asset (Vite, webpack 5 — the same `new URL(..., import.meta.url)` pattern webpack 5 documents); no plugins, no `?url` imports, no copy step. One exception: **Vite's dev server** does not rewrite the expression inside pre-bundled dependencies — see the caveat below.
2. **Node** — the binary is read from disk next to the installed package and initialized synchronously (see [Node / SSR](#node-ssr)).

### Optional: `configureWasm`

An escape hatch for setups where the defaults cannot work — self-hosted copies on a CDN, Service Worker environments, bundlers without asset-URL support, or load-timeout tuning:

```ts
import { configureWasm } from "@marcusok/excel-exporter";

configureWasm({
  // Both optional; set only what you want to override.
  wasmUrl: "https://cdn.example.com/modern-xlsx.wasm",
  workerUrl: "https://cdn.example.com/export.worker.js",
});
```

Bundlers with asset imports can wire the shipped files explicitly instead (fully supported, the pre-2.0 recommended setup):

```ts
import wasmUrl from "@marcusok/excel-exporter/dist/modern-xlsx.wasm?url";
import workerUrl from "@marcusok/excel-exporter/dist/export.worker.js?url";
configureWasm({ wasmUrl, workerUrl });
```

| Option            | Type            | Default                 | Description                                                            |
| ----------------- | --------------- | ----------------------- | ---------------------------------------------------------------------- |
| `wasmUrl`         | `string \| URL` | the shipped `.wasm`     | Override for a self-hosted / CDN copy                                  |
| `workerUrl`       | `string \| URL` | the shipped worker file | Override for a self-hosted / CDN copy                                  |
| `timeoutMs`       | `number`        | `10_000`                | Per-attempt load timeout                                               |
| `maxRetries`      | `number`        | `3`                     | Max attempts; failed attempts wait 300ms, then 600ms                   |
| `workerTimeoutMs` | `number`        | `120_000`               | Worker export timeout; a timed-out export terminates the shared worker |

`configureWasm` merges options: only a changed `wasmUrl` resets an already-loaded (or mid-load) WASM instance; changing timeouts/retries alone never causes re-initialization. If a previous load failed (error state), any `configureWasm` call clears the error so the next export retries with the new settings.

### Vite dev server: pre-bundling caveat

Vite's dev server pre-bundles dependencies with esbuild into `/node_modules/.vite/deps/`. Inside the pre-bundled chunk, `import.meta.url` points into `.vite/deps/`, so `new URL("./modern-xlsx.wasm", import.meta.url)` resolves to `/node_modules/.vite/deps/modern-xlsx.wasm` — a path where the file does not exist. Vite's HTML fallback (active for `Accept: text/html` **and** `Accept: */*` — plain `fetch` sends the latter) answers that request with `index.html` (HTTP 200, `text/html`; under `appType: 'mpa'` you get a 404 instead — same cause, same fix), WASM compilation fails on the HTML bytes (`expected magic word 00 61 73 6d, found 3c 21 64 6f` — `3c 21 64 6f` is `<!do` from `<!doctype html>`), and the export **silently degrades to the style-less stream**: the file still downloads and `result.success` is `true`, but styles, column widths, freeze panes and auto-filters are all stripped, and `result.error` mentions `Fallback: styles stripped (fast stream)`.

Scope: **`vite build` is unaffected** — the production pipeline emits the wasm as a hashed asset correctly; only the dev server with default `optimizeDeps` is affected. The worker asset (`export.worker.js`, auto mode ≥ 20,000 rows) resolves through the same mechanism and fails the same way, so this is not limited to small exports. One diagnostic trap: after the first failed attempt, later exports report `WASM load previously failed` instead of the original error — the real reason only appears in the console warning of the **first** export (or after a page reload). One exception: on a slow network the underlying load can still succeed _after_ all retry attempts have timed out; the loader then heals itself back to ready, and later exports run the fully-styled paths again without any `configureWasm` call (the "previously failed" message stops appearing).

Fix — exclude the package from pre-bundling in `vite.config.ts`:

```ts
import { defineConfig } from "vite";

export default defineConfig({
  optimizeDeps: {
    exclude: ["@marcusok/excel-exporter"],
  },
});
```

Restart the dev server afterwards (the `.vite/deps` cache must be re-created). With the package excluded, Vite serves its ESM output directly from `node_modules`, `import.meta.url` resolves back to the real `dist/` location, and both assets load as shipped. The alternative — `?url` imports wired through `configureWasm` (above) — also works without touching `optimizeDeps`, at the cost of two lines in your entry module.

## Node / SSR

No browser static assets and **no initialization boilerplate** are needed in Node: with nothing configured, the engine locates this package's `dist/modern-xlsx.wasm` on disk (pnpm-symlink-safe) and initializes it synchronously on first use. `auto` never uses Workers in Node; ≥ 50k rows switch to streaming on the main thread (the stream path does not use WASM). See [Node/SSR](/packages/excel-exporter/guide/09-node-ssr) for explicit-init timing control and bundler caveats.
