import { initWasm, initWasmSync } from "modern-xlsx";

export type LoadState = "idle" | "loading" | "ready" | "error";

export interface LoaderOptions {
  /**
   * WASM URL. Defaults to the binary shipped next to this package's entry
   * (`dist/modern-xlsx.wasm`): bundlers that support the `new URL(asset,
   * import.meta.url)` pattern (Vite, webpack 5, Rollup) rewrite it to a hashed
   * asset automatically, and Node locates it through `node_modules`. Override
   * only for self-hosted copies, a CDN, or bundlers without asset-URL support.
   */
  wasmUrl?: string | URL;
  /**
   * export.worker.js URL, required for worker mode. Defaults to the
   * self-contained worker shipped next to this package's entry — the same
   * bundler rewrite applies. Override for self-hosted copies.
   */
  workerUrl?: string | URL;
  /** Per-attempt load timeout, default 10s. */
  timeoutMs?: number;
  /** Max load attempts (total, including the first), default 3. */
  maxRetries?: number;
  /**
   * Worker export timeout, default 120s. A timed-out export terminates the
   * shared worker and rejects its sibling requests, so raise this only for
   * legitimately huge exports (and prefer splitting into multiple sheets).
   */
  workerTimeoutMs?: number;
}

/**
 * Default WASM location: the binary this package ships next to its entry.
 * Kept as a `new URL(<literal>, import.meta.url)` expression — the exact form
 * Vite's `vite:asset-import-meta-url` / webpack 5 match to emit a hashed
 * asset at build time, and the natural relative location in Node.
 */
export function defaultWasmUrl(): URL {
  return new URL("./modern-xlsx.wasm", import.meta.url);
}

export class WasmLoader {
  private state: LoadState = "idle";
  private promise: Promise<void> | null = null;
  private opts: LoaderOptions;

  constructor(opts: LoaderOptions = {}) {
    this.opts = { timeoutMs: 10_000, maxRetries: 3, ...opts };
  }

  get supported(): boolean {
    return (
      typeof WebAssembly !== "undefined" &&
      typeof WebAssembly.instantiate === "function"
    );
  }

  get isReady(): boolean {
    return this.state === "ready";
  }

  getOptions(): Readonly<LoaderOptions> {
    return this.opts;
  }

  /**
   * Merge new options into the current set. If the WASM URL changes while the
   * loader is already ready (or mid-load), reset so the next ensureLoaded
   * re-attempts initialization with the new URL; otherwise keep the loaded
   * state. This avoids discarding an already-loaded WASM module when only
   * timeouts/retries change. A previous load *error* is always cleared by a
   * reconfiguration, so the next ensureLoaded retries with the new settings
   * instead of throwing forever.
   *
   * Caveat (modern-xlsx 1.2.0 verified): `initWasm` is idempotent with a
   * module-level "first successful init wins" guard. On a thread where WASM
   * is already initialized the re-attempt is a silent no-op — the reset only
   * guarantees initWasm is *called* with the new URL, which modern-xlsx
   * ignores if its module-level `initialized` flag is already set. The new
   * URL genuinely takes effect only in a fresh JS realm (a page reload, or a
   * worker created after terminateWorker()). updateOptions warns when this
   * caveat applies.
   */
  updateOptions(opts: LoaderOptions): void {
    const urlChanged =
      opts.wasmUrl !== undefined && opts.wasmUrl !== this.opts.wasmUrl;
    if (urlChanged && this.state === "ready") {
      console.warn(
        "[excel-exporter] wasmUrl changed after WASM already initialized on this thread. " +
          "modern-xlsx's initWasm is idempotent (first successful init wins), so the already-loaded " +
          "module stays in effect and the new URL is ignored by initWasm. The new URL takes effect " +
          "only in a fresh JS realm (reload the page, or terminateWorker() before the next export " +
          "so a new worker is created).",
      );
    }
    this.opts = { ...this.opts, ...opts };
    if ((urlChanged && this.state !== "idle") || this.state === "error") {
      this.state = "idle";
      this.promise = null;
    }
  }

  async ensureLoaded(): Promise<void> {
    if (this.state === "ready") return;
    if (this.state === "error") {
      throw new Error(
        "[excel-exporter] WASM load previously failed; call configureWasm() to retry with new settings",
      );
    }
    if (this.promise) return this.promise;
    // Capture the promise locally: updateOptions() may null this.promise while
    // the load is in flight (wasmUrl changed), and this load must not clobber
    // the reset state when it settles -- otherwise a superseded old-URL load
    // would mark the loader ready and the new URL would never take effect.
    const promise = (this.promise = this.loadWithRetry());
    try {
      await promise;
      if (this.promise === promise) this.state = "ready";
    } catch (e) {
      if (this.promise === promise) this.state = "error";
      throw e;
    }
  }

  /**
   * Node auto-init: when no wasmUrl is configured and we run on Node, read
   * this package's own `dist/modern-xlsx.wasm` from disk and initialize
   * synchronously (Node's fetch rejects the file:// URL the browser default
   * would produce). This removes any init boilerplate from Node consumers —
   * nothing to call, nothing to copy.
   *
   * Resolution order within this method:
   *  1. `./modern-xlsx.wasm` next to the published entry (dist/).
   *  2. `../dist/modern-xlsx.wasm` — the binary's location when this module
   *     runs from src/ (repo tests, source-aliased monorepo consumers).
   *
   * Returns false in browsers, when a URL is configured, or on any failure,
   * so the standard initWasm path (and its stream degradation) is untouched.
   *
   * Node built-ins MUST stay dynamically imported: this module also ships in
   * browser bundles, where a static `import "node:fs"` would fail to resolve.
   */
  private async tryNodeAutoInit(): Promise<boolean> {
    if (this.opts.wasmUrl !== undefined) return false;
    if (typeof process === "undefined" || !process.versions?.node) return false;
    try {
      // Test suites mock modern-xlsx with a bare { initWasm } factory; a
      // missing initWasmSync must skip auto-init, not throw a TypeError.
      if (typeof initWasmSync !== "function") return false;
      const fsNs = await import("node:fs");
      let bytes: Uint8Array;
      try {
        bytes = fsNs.readFileSync(defaultWasmUrl());
      } catch {
        bytes = fsNs.readFileSync(
          new URL("../dist/modern-xlsx.wasm", import.meta.url),
        );
      }
      initWasmSync(bytes);
      return true;
    } catch {
      // Resolution/read/init failure (e.g. a consumer bundling for Node
      // without the package files on disk): fall through to initWasm,
      // which retries and degrades exactly as before this path existed.
      return false;
    }
  }

  private async loadWithRetry(): Promise<void> {
    if (!this.supported) {
      throw new Error(
        "[excel-exporter] WebAssembly not supported in this environment",
      );
    }
    // Node without a configured URL: locate and init the wasm synchronously
    // (initWasmSync) instead of fetching a file:// URL that Node's fetch
    // rejects. Idempotent with initWasm (shared "initialized" flag inside
    // modern-xlsx), so a later initWasm call on the same thread is a no-op.
    if (await this.tryNodeAutoInit()) return;
    const wasmUrl = this.opts.wasmUrl ?? defaultWasmUrl();
    const timeoutMs = this.opts.timeoutMs ?? 10_000;
    const maxRetries = this.opts.maxRetries ?? 3;
    let lastErr: unknown;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      const timeout = new Promise<never>((_, reject) => {
        timer = setTimeout(
          () => reject(new Error(`WASM load timeout (attempt ${attempt})`)),
          timeoutMs,
        );
      });
      try {
        this.state = "loading";
        await Promise.race([initWasm(wasmUrl), timeout]);
        return;
      } catch (e) {
        lastErr = e;
        if (attempt < maxRetries) {
          await new Promise((r) => setTimeout(r, 300 * 2 ** (attempt - 1)));
        }
      } finally {
        // Clear the pending timeout so a late reject never surfaces as an
        // unhandled promise rejection after initWasm already resolved.
        if (timer) clearTimeout(timer);
      }
    }
    throw new Error(
      `[excel-exporter] WASM load failed after ${maxRetries} attempts: ${(lastErr as Error).message}`,
    );
  }
}

const defaultLoader: WasmLoader = new WasmLoader();

export function getWasmLoader(): WasmLoader {
  return defaultLoader;
}

/**
 * Configure WASM / worker URLs and timeout settings. Entirely optional since
 * the assets default to their shipped locations (see LoaderOptions): call
 * this only to point at self-hosted copies, a CDN, or a custom build.
 *
 * Merges into the existing loader rather than replacing it, so an
 * already-loaded WASM module is kept unless the WASM URL actually changes. A
 * previous load error is always cleared, so calling this after a failure
 * makes the next export retry with the new settings.
 *
 * Note: changing `wasmUrl` after a *successful* load does not reload WASM on
 * a thread that already initialized it — modern-xlsx's `initWasm` is
 * idempotent and keeps the first successfully loaded module (see
 * WasmLoader.updateOptions). The new URL takes effect in a fresh JS realm
 * only (page reload / a worker created after `terminateWorker()`), and
 * updateOptions prints a warning when the caveat applies.
 */
export function configureWasm(opts: LoaderOptions): void {
  defaultLoader.updateOptions(opts);
}
