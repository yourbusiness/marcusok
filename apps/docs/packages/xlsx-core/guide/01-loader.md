# WASM Loader

The loader is the reason this package exists. It locates the WASM binary, initializes it exactly once per JS realm and reports failures through a typed state machine — with zero configuration in both the browser and Node.

```ts
import { getWasmLoader, readBuffer } from "@marcusok/xlsx-core";

// Optional: only for self-hosted copies / a CDN / a custom build.
// configureWasm({ wasmUrl: "/assets/modern-xlsx.wasm" });

await getWasmLoader().ensureLoaded(); // no-op if already ready
const workbook = await readBuffer(bytes);
```

Most consumers never call `ensureLoaded()` explicitly: the business packages await it on their first operation, so the engine is initialized lazily. Call it yourself when you want the one-off read + compile charged to a known moment (process startup, a loading screen) instead of to the first user action.

## How assets resolve by default

| Environment | What happens                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| ----------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Browser     | `defaultWasmUrl()` returns `new URL("./modern-xlsx.wasm", import.meta.url)` — the form Vite's `vite:asset-import-meta-url` and webpack 5 statically rewrite into a hashed asset emitted next to your bundle. The loader then fetches that URL and calls the engine's `initWasm(wasmUrl)`.                                                                                                                                                                                                                                                                                                |
| Node        | With nothing configured, the loader reads its own `dist/modern-xlsx.wasm` from disk and calls `initWasmSync(bytes)` — Node's `fetch` rejects the `file://` URL the browser default would produce, so the sync path is the only correct one. `node:fs` is imported dynamically with a computed specifier, keeping the module bundler-safe for browser targets. The import chain: `./modern-xlsx.wasm` next to the published entry (the normal case) → `../dist/modern-xlsx.wasm` (the location when this module runs from `src/`, i.e. repo tests and source-aliased monorepo consumers). |

Configure an explicit `wasmUrl` and the Node auto-init is skipped — you asked for that URL, so the loader fetches it through `initWasm` like a browser would. Any resolution/read/init failure inside auto-init returns `false` silently and falls through to the normal `initWasm` path, so a deployment that bundles the code without the asset degrades exactly as it did before the auto-init existed.

## LoaderOptions

Every field is optional. `configureWasm(options)` merges into the current set.

| Option            | Type            | Default                                                 | Meaning                                                                                                                                                                                                                                                                                                                                               |
| ----------------- | --------------- | ------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `wasmUrl`         | `string \| URL` | the shipped `dist/modern-xlsx.wasm`                     | Override for a self-hosted copy, a CDN or a bundler without asset-URL support                                                                                                                                                                                                                                                                         |
| `workerUrl`       | `string \| URL` | the export worker shipped by `@marcusok/excel-exporter` | Export-worker script. Read by the exporter only — one field cannot serve two packages (the preview would spawn the export worker), hence the separate option below                                                                                                                                                                                    |
| `parseWorkerUrl`  | `string \| URL` | the parse worker shipped by `@marcusok/excel-preview`   | Parse-worker script. Read by the preview only                                                                                                                                                                                                                                                                                                         |
| `timeoutMs`       | `number`        | `10_000`                                                | Per-attempt load timeout                                                                                                                                                                                                                                                                                                                              |
| `maxRetries`      | `number`        | `3`                                                     | Total load attempts, including the first. Failed attempts back off 300ms, then 600ms                                                                                                                                                                                                                                                                  |
| `workerTimeoutMs` | `number`        | `120_000`                                               | Worker operation timeout, consumed by the two business packages. A timed-out operation terminates the shared worker and rejects its sibling requests, so raise it only for legitimately huge workloads. Must be `> 0`: both packages pass the value straight to `setTimeout` (`value ?? 120_000`), so `0` means "time out immediately", not "disable" |

Defaults come from two places, which is worth knowing when you read the table. `timeoutMs` / `maxRetries` are applied by `new WasmLoader()` itself (the constructor seeds `{ timeoutMs: 10_000, maxRetries: 3 }`), and `wasmUrl` falls back to `defaultWasmUrl()` at load time (`this.opts.wasmUrl ?? defaultWasmUrl()`). The worker options are the opposite: the loader never resolves them, it only stores what you pass, and each consuming package picks its own shipped default — `workerUrl ?? new URL("./export.worker.js", import.meta.url)` in the exporter, `parseWorkerUrl ?? new URL("./parse.worker.js", import.meta.url)` in the preview — and reads `getOptions().workerTimeoutMs ?? 120_000` at call time. A page that consumes this package directly therefore has no worker script at all: the worker options exist for the two business packages.

## Timeout and retry semantics

An attempt is one `initWasm(wasmUrl)` call raced against a `timeoutMs` timer. On failure the next attempt starts after a `300 * 2 ** (attempt - 1)` ms backoff — 300ms, then 600ms for the default three attempts. When all attempts fail, `ensureLoaded()` rejects with:

```
[xlsx-core] WASM load failed after 3 attempts: <last error message>
```

Two failure modes are worth distinguishing:

- **No WebAssembly in the environment.** `loadWithRetry` rejects immediately with `[xlsx-core] WebAssembly not supported in this environment`, without consuming retries. Check `getWasmLoader().supported` (or `typeof WebAssembly`) before offering WASM-dependent UI.
- **Every attempt timed out, but the load later succeeds anyway.** modern-xlsx's `initWasm` keeps a single module-level in-flight promise, so a slow network can still resolve the original fetch after the loader gave up. The loader watches for that and heals itself: a late success flips `error` back to `ready` and clears the pending promise, after which the next operation works without any `configureWasm()` call. Without this hook the `error` state would permanently contradict the engine's real state. The same mechanism explains the diagnostic trap on the [exporter's installation page](/packages/excel-exporter/guide/02-installation#vite-dev-server-pre-bundling-caveat): after a first failure, later attempts report `WASM load previously failed` instead of the original cause.

## The load state machine

`LoadState` is `"idle" | "loading" | "ready" | "error"`, and it is **private state**: the field is not part of the public class, so read it through the two accessors and through the outcome of `ensureLoaded()`.

| State     | How you get there                                           | What `ensureLoaded()` does                                                                                                    |
| --------- | ----------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| `idle`    | Initial state; also after `configureWasm` resets the loader | Sets `loading` and starts the retry loop                                                                                      |
| `loading` | A load is in flight                                         | Returns the **same** promise — concurrent callers single-flight onto one load instead of starting a second one                |
| `ready`   | All attempts done (or a late success self-healed it)        | Returns immediately; `loader.isReady` is `true`                                                                               |
| `error`   | All attempts failed                                         | Throws `[xlsx-core] WASM load previously failed; call configureWasm() to retry with new settings` without retrying on its own |

When to query it: use `isReady` for questions that must **not** trigger work ("can I show the styled routes?"), and `ensureLoaded()` for questions that should ("give me the engine"). The state is written only inside `ensureLoaded()`, guarded by promise identity, so a superseded load (URL changed mid-flight) can neither mark the loader ready nor hide a newer load's result.

`configureWasm()` clears an `error` state unconditionally, while a **change** to `wasmUrl` also resets `ready` / `loading` back to `idle`. Changing only timeouts/retries keeps an already-loaded module — there is no reason to recompile WASM because a timeout changed.

## `WasmLoader` and `defaultWasmUrl()`

- `WasmLoader` is the class behind the singleton. `getWasmLoader()` returns the module-level instance; constructing your own (`new WasmLoader({ timeoutMs: 1000 })`) is supported (tests, isolation of loader state) but gives you **loader state isolation only** — modern-xlsx's `initWasm` is module-global, so a second loader cannot produce a second engine instance on the same thread.
- `defaultWasmUrl()` returns the shipped binary's URL as a `URL` object. It is public because consumers need to _forward_ the resolved default: a worker script that was copied or renamed cannot reliably resolve `./modern-xlsx.wasm` relatively, so the business packages pass `String(wasmUrl ?? defaultWasmUrl())` over `postMessage` instead (a `URL` object is not structured-cloneable and throws `DataCloneError`).

## Self-hosting and CDN overrides

```ts
import { configureWasm } from "@marcusok/xlsx-core";

configureWasm({ wasmUrl: "https://cdn.example.com/modern-xlsx.wasm" });
```

Or wire the shipped file through your bundler's asset imports:

```ts
import wasmUrl from "@marcusok/xlsx-core/dist/modern-xlsx.wasm?url";
configureWasm({ wasmUrl });
```

**Call it before the first load.** modern-xlsx's `initWasm` is idempotent with a "first successful init wins" guard and holds one in-flight promise, so changing `wasmUrl` after a successful load does not reload anything, and changing it while the initial fetch is pending cannot redirect that fetch. The new URL takes effect in a fresh JS realm only — a page reload, or a worker created after the shared worker was terminated. `updateOptions()` prints a console warning whenever that caveat applies (the same applies to `workerUrl` / `parseWorkerUrl`: the shared workers read their script URL once at creation).

Worker scripts are a second constraint: a browser rejects a cross-origin worker script at `Worker` construction. Only the WASM binary may come from a plain CDN. The two business packages document their own overrides in [Use the loader from another package](#use-the-loader-from-another-package) below and in the [exporter's installation guide](/packages/excel-exporter/guide/02-installation#optional-configurewasm) / [preview's assets guide](/packages/excel-preview/guide/02-assets).

## Vite dev server: pre-bundling caveat

The same caveat the exporter documents applies to direct consumers of this package, with the same cause and the same fix: Vite's dev server pre-bundles dependencies into `node_modules/.vite/deps/`, where `import.meta.url` no longer points at the real `dist/`, so the default WASM URL resolves to a path that does not exist and the request comes back as `index.html`. Add the package that provides the loader to `optimizeDeps.exclude`:

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ["@marcusok/xlsx-core"] },
});
```

`vite build` is unaffected (the production pipeline emits the wasm as a hashed asset correctly), and the `?url` + `configureWasm` wiring above also works without touching `optimizeDeps`. The full diagnosis — including the `expected magic word` compile error and the silent-degradation symptoms the business packages show — is in the [exporter's installation notes](/packages/excel-exporter/guide/02-installation#vite-dev-server-pre-bundling-caveat); Node servers have their own form of it, described in the [Node/SSR guide](/packages/excel-exporter/guide/09-node-ssr).

## Use the loader from another package

`configureWasm` and `getWasmLoader` are re-exported by both business packages, and all three names point at this package's single loader:

```ts
import { configureWasm, getWasmLoader } from "@marcusok/excel-exporter"; // or @marcusok/excel-preview
```

That is the recommended entry point when you already use one of them — the option semantics, defaults and warnings documented here are exactly what you get, and the [preview's API reference](/packages/excel-preview/api/01-create-preview#other-exports) lists them in its own surface for the same reason.
