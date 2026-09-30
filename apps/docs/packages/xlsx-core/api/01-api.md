# Loader API

The complete surface of `@marcusok/xlsx-core`'s own code — the loader. The rest of the package's exports are modern-xlsx re-exports, listed in the [engine surface](/packages/xlsx-core/guide/02-engine-surface).

```ts
import {
  WasmLoader,
  configureWasm,
  getWasmLoader,
  defaultWasmUrl,
  type LoaderOptions,
  type LoadState,
} from "@marcusok/xlsx-core";
```

## `configureWasm(options: LoaderOptions): void`

Merges `options` into the shared loader's current set. Entirely optional: with nothing configured, assets resolve to their shipped locations.

- A previous load **error** is always cleared, so calling this after a failure makes the next `ensureLoaded()` retry with the new settings instead of throwing "previously failed" forever.
- Changing only `timeoutMs` / `maxRetries` / `workerTimeoutMs` keeps an already-loaded WASM module.
- Changing `wasmUrl` resets a `ready` / `loading` loader to `idle`. This guarantees `initWasm` is _called_ with the new URL — it does not guarantee the new URL is _used_: modern-xlsx's `initWasm` is idempotent and keeps a single in-flight promise, so on a thread that already initialized WASM the call is a silent no-op, and a still-pending initial fetch cannot be aborted or redirected. The URL genuinely takes effect only in a fresh JS realm (page reload, or a worker created after the shared worker was terminated). `updateOptions` prints a console warning when the caveat applies.
- Changing `workerUrl` / `parseWorkerUrl` always warns: the shared workers read their script URL once at creation and are reused afterwards.

## `getWasmLoader(): WasmLoader`

Returns the module-level singleton. Every importer of every @marcusok package gets the same object.

## `defaultWasmUrl(): URL`

The shipped binary's location: `new URL("./modern-xlsx.wasm", import.meta.url)`, kept as a literal `new URL` expression so bundlers emit it as an asset. Use it to forward the resolved default somewhere that cannot resolve a relative URL — a worker gets `String(wasmUrl ?? defaultWasmUrl())` over `postMessage`, because a `URL` object is not structured-cloneable.

## `class WasmLoader`

| Member                | Signature                       | Notes                                                                                         |
| --------------------- | ------------------------------- | --------------------------------------------------------------------------------------------- |
| `constructor`         | `(opts?: LoaderOptions)`        | Seeds `{ timeoutMs: 10_000, maxRetries: 3 }` and merges the rest                              |
| `supported` (getter)  | `boolean`                       | `typeof WebAssembly !== "undefined" && typeof WebAssembly.instantiate === "function"`         |
| `isReady` (getter)    | `boolean`                       | `state === "ready"` — the only public view of the state machine, and it never triggers a load |
| `getOptions()`        | `Readonly<LoaderOptions>`       | The current merged options (what the consuming packages read for `workerTimeoutMs`)           |
| `updateOptions(opts)` | `(opts: LoaderOptions) => void` | What `configureWasm` calls on the singleton                                                   |
| `ensureLoaded()`      | `() => Promise<void>`           | Single-flight; resolves when the engine is usable, rejects with the load error otherwise      |

Constructing your own instance is supported, but it isolates **loader** state only — modern-xlsx's initialization is module-global, so a second instance cannot produce a second engine on the same thread.

`ensureLoaded()` resolves immediately when ready, joins the in-flight promise when loading, starts a load when idle, and rejects with `[xlsx-core] WASM load previously failed; call configureWasm() to retry with new settings` when in error state. A failed load rejects with `[xlsx-core] WASM load failed after <n> attempts: <last error>`; an environment without WebAssembly rejects with `[xlsx-core] WebAssembly not supported in this environment` before any attempt.

## `LoaderOptions`

| Field             | Type            | Default                                   | Consumed by                | Meaning                                                                                                                                        |
| ----------------- | --------------- | ----------------------------------------- | -------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| `wasmUrl`         | `string \| URL` | `defaultWasmUrl()` (the shipped binary)   | the loader                 | Explicit WASM location; also disables the Node on-disk auto-init                                                                               |
| `workerUrl`       | `string \| URL` | the exporter's shipped `export.worker.js` | `@marcusok/excel-exporter` | Export-worker script URL                                                                                                                       |
| `parseWorkerUrl`  | `string \| URL` | the preview's shipped `parse.worker.js`   | `@marcusok/excel-preview`  | Parse-worker script URL — a separate field so a page using both packages cannot cross-wire them                                                |
| `timeoutMs`       | `number`        | `10_000`                                  | the loader                 | Per-attempt load timeout                                                                                                                       |
| `maxRetries`      | `number`        | `3`                                       | the loader                 | Total attempts including the first; backoff `300 * 2 ** (attempt - 1)` ms                                                                      |
| `workerTimeoutMs` | `number`        | `120_000`                                 | both business packages     | Worker operation timeout; a timeout terminates the shared worker and rejects its sibling requests. Must be `> 0` — there is no "disable" value |

All fields are optional. Only the loader-side defaults are stored in the loader; the worker defaults are resolved by the consuming package (see the [loader guide](/packages/xlsx-core/guide/01-loader#loaderoptions)).

## `LoadState`

```ts
type LoadState = "idle" | "loading" | "ready" | "error";
```

| Value     | Meaning                                                              | `ensureLoaded()` behaviour                       |
| --------- | -------------------------------------------------------------------- | ------------------------------------------------ |
| `idle`    | Nothing started, or the loader was reset by a `wasmUrl` change       | Starts the load                                  |
| `loading` | A load is in flight                                                  | Joins the same promise                           |
| `ready`   | The engine is initialized on this thread                             | Resolves immediately                             |
| `error`   | All attempts failed (a late success self-heals this back to `ready`) | Rejects until `configureWasm()` clears the state |

The state field itself is private: observe it through `isReady` and the outcome of `ensureLoaded()`.

## `@marcusok/xlsx-core/tsup`

```ts
import {
  rewriteWasmBgUrl,
  dropNodeFsPromises,
  type EsbuildPlugin,
} from "@marcusok/xlsx-core/tsup";

declare const rewriteWasmBgUrl: EsbuildPlugin;
declare const dropNodeFsPromises: EsbuildPlugin;
```

`EsbuildPlugin` is `NonNullable<Options["esbuildPlugins"]>[number]` — derived from tsup's own option shape, so the plugins drop straight into `esbuildPlugins`. Node-only (imports `node:fs`); see [package integration](/packages/xlsx-core/guide/03-package-integration).

## Complete example

```ts
import {
  configureWasm,
  defaultWasmUrl,
  getWasmLoader,
  readBuffer,
  formatCellRich,
  ModernXlsxError,
  WASM_INIT_FAILED,
  type LoaderOptions,
} from "@marcusok/xlsx-core";

// 1. Optional: self-hosted assets and/or a longer first-load timeout.
configureWasm({
  wasmUrl: "/assets/modern-xlsx.wasm",
  timeoutMs: 15_000,
} satisfies LoaderOptions);

// 2. Optional: warm the engine up front (Node reads it from disk synchronously).
const loader = getWasmLoader();
if (loader.supported && !loader.isReady) {
  await loader.ensureLoaded(); // throws on failure — configureWasm() re-arms it
}

// 3. Use the engine.
try {
  const workbook = await readBuffer(bytes);
  console.log(workbook.sheetNames, formatCellRich(1234.5, "#,##0.00").text);
} catch (err) {
  if (err instanceof ModernXlsxError && err.code === WASM_INIT_FAILED) {
    console.error(
      "engine unavailable, resolved wasm URL was",
      String(defaultWasmUrl()),
    );
  }
}
```
