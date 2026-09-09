import { initWasm } from "modern-xlsx";

export type LoadState = "idle" | "loading" | "ready" | "error";

export interface LoaderOptions {
  /** Self-hosted WASM URL. Strongly recommended in production to avoid CDN drift. */
  wasmUrl?: string | URL;
  /** Self-hosted export.worker.js URL, required for worker mode. */
  workerUrl?: string | URL;
  /** Per-attempt load timeout, default 10s. */
  timeoutMs?: number;
  /** Max load attempts (total, including the first), default 3. */
  maxRetries?: number;
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
   * Node auto-init: when no wasmUrl is configured and we run on Node, locate
   * modern-xlsx's wasm binary via createRequire (the same resolution the docs
   * recommend for manual initWasmSync) and initialize synchronously. This
   * removes the boilerplate initWasmSync(readFileSync(...)) from every Node
   * consumer's entry file.
   *
   * Returns false in browsers, when a URL is configured, or on any failure,
   * so the standard initWasm path (and its SheetJS degradation) is untouched.
   *
   * Node built-ins MUST stay dynamically imported: this module also ships in
   * browser bundles, where a static `import "node:fs"` would fail to resolve.
   */
  private async tryNodeAutoInit(): Promise<boolean> {
    if (this.opts.wasmUrl !== undefined) return false;
    if (typeof process === "undefined" || !process.versions?.node) return false;
    try {
      const [moduleNs, fsNs, pathNs, mx] = await Promise.all([
        import("node:module"),
        import("node:fs"),
        import("node:path"),
        import("modern-xlsx"),
      ]);
      // Test suites mock modern-xlsx with a bare { initWasm } factory; a
      // missing initWasmSync must skip auto-init, not throw a TypeError.
      if (typeof mx.initWasmSync !== "function") return false;
      const require = moduleNs.createRequire(import.meta.url);
      const wasmPath = `${pathNs.dirname(require.resolve("modern-xlsx"))}/modern-xlsx.wasm`;
      mx.initWasmSync(fsNs.readFileSync(wasmPath));
      return true;
    } catch {
      // Resolution/read/init failure (e.g. a consumer bundling for Node
      // without the runtime package on disk): fall through to initWasm,
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
    const wasmUrl = this.opts.wasmUrl;
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
 * Inject CDN / self-hosted URLs and timeout config at app entry. Merges into the
 * existing loader rather than replacing it, so an already-loaded WASM module is
 * kept unless the WASM URL actually changes. A previous load error is always
 * cleared, so calling this after a failure makes the next export retry with the
 * new settings.
 *
 * Note: changing `wasmUrl` after a *successful* load does not reload WASM on a
 * thread that already initialized it — modern-xlsx's `initWasm` is idempotent
 * and keeps the first successfully loaded module (see WasmLoader.updateOptions).
 * The new URL takes effect in a fresh JS realm only (page reload / a worker
 * created after `terminateWorker()`), and updateOptions prints a warning when
 * the caveat applies.
 */
export function configureWasm(opts: LoaderOptions): void {
  defaultLoader.updateOptions(opts);
}
