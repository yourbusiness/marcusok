# Installation & Configuration

## Requirements

- Node `>= 22` (the package's `engines` requirement); any package manager works — examples use pnpm, npm / yarn are equivalent
- Browsers need WebAssembly support (all modern browsers)

## Install

```bash
pnpm add @marcusok/excel-exporter
```

That is the entire setup. The package has **zero runtime dependencies**: the export engine (modern-xlsx JS glue + fflate) is bundled in at build time, and the WASM binary ships under this package's own `exports` map — there is no engine package to install, no optional fallback package, and nothing to add to `main.ts` or your bundler config.

## Browser: assets resolve automatically

Two files ship alongside the code and are located by default, so a plain `import { exportExcel } from "@marcusok/excel-exporter"` works out of the box:

| Asset              | Description                                                                                                                                                 |
| ------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `modern-xlsx.wasm` | WASM core (~1.9MB); needed on every route except an explicit `mode: "stream"` (pure JS)                                                                     |
| `export.worker.js` | Self-contained worker entry (a single ESM file, zero imports); needed only when exports enter a Worker (auto ≥ 20,000 rows, or explicit worker/stream mode) |

Resolution order:

1. **Bundlers** — both default to `new URL(<file>, import.meta.url)` next to the package entry. Vite rewrites this in dev (dependency pre-bundling) and emits hashed assets at build time; webpack 5 documents the same asset pattern. No plugins, no `?url` imports, no copy step.
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

| Option       | Type            | Default                 | Description                                          |
| ------------ | --------------- | ----------------------- | ---------------------------------------------------- |
| `wasmUrl`    | `string \| URL` | the shipped `.wasm`     | Override for a self-hosted / CDN copy                |
| `workerUrl`  | `string \| URL` | the shipped worker file | Override for a self-hosted / CDN copy                |
| `timeoutMs`  | `number`        | `10_000`                | Per-attempt load timeout                             |
| `maxRetries` | `number`        | `3`                     | Max attempts; failed attempts wait 300ms, then 600ms |

`configureWasm` merges options: only a changed `wasmUrl` resets an already-loaded (or mid-load) WASM instance; changing timeouts/retries alone never causes re-initialization. If a previous load failed (error state), any `configureWasm` call clears the error so the next export retries with the new settings.

## Node / SSR

No browser static assets and **no initialization boilerplate** are needed in Node: with nothing configured, the engine locates this package's `dist/modern-xlsx.wasm` on disk (pnpm-symlink-safe) and initializes it synchronously on first use. `auto` never uses Workers in Node; ≥ 50k rows switch to streaming on the main thread (the stream path does not use WASM). See [Node/SSR](/packages/excel-exporter/guide/09-node-ssr) for explicit-init timing control and bundler caveats.
