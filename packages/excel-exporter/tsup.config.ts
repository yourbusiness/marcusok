import fs from "node:fs";
import { defineConfig, type Options } from "tsup";

/** esbuild plugin type, derived from tsup's own option shape (esbuild itself is a transitive dep under pnpm). */
type EsbuildPlugin = NonNullable<Options["esbuildPlugins"]>[number];

/**
 * modern-xlsx's wasm-bindgen glue (wasm/modern_xlsx_wasm.js) defaults to
 * `new URL('modern_xlsx_wasm_bg.wasm', import.meta.url)` — a filename this
 * package does not ship (we forward the identical binary as
 * dist/modern-xlsx.wasm; sha256-verified equal). Our code always passes an
 * explicit URL, so the branch is dead at runtime, but consumer bundlers
 * statically analyze every `new URL(<literal>, import.meta.url)` and warn
 * when the file is missing (Vite: "doesn't exist at build time, it will
 * remain unchanged...").
 *
 * The replacement is deliberately NOT a string literal (array join): neither
 * esbuild nor downstream bundlers match computed specifiers, so nothing is
 * asset-analyzed or emitted twice. If the branch ever executes at runtime it
 * still resolves correctly — ["..","dist","modern-xlsx.wasm"].join("/") from
 * the glue file and from this package's dist both point at the shipped copy.
 */
const rewriteWasmBgUrl: EsbuildPlugin = {
  name: "rewrite-modern-xlsx-wasm-bg-url",
  setup(build) {
    build.onLoad({ filter: /modern_xlsx_wasm\.js$/ }, async (args) => {
      const contents = await fs.promises.readFile(args.path, "utf8");
      if (!contents.includes("modern_xlsx_wasm_bg.wasm")) return undefined;
      return {
        contents: contents.replace(
          /new URL\((['"])modern_xlsx_wasm_bg\.wasm\1,\s*import\.meta\.url\)/g,
          'new URL(["..", "dist", "modern-xlsx.wasm"].join("/"), import.meta.url)',
        ),
        loader: "js",
      };
    });
  },
};

/**
 * modern-xlsx's Node-only file APIs (`toFile`, `readFile`) dynamic-import
 * `node:fs/promises`. This package never exports those APIs (only buffer
 * APIs: toBuffer/readBuffer), so in the bundled output the imports are dead
 * branches — but they survive as literal dynamic imports, and consumer
 * browser builds warn about them (Vite 5/VitePress: "Module fs/promises has
 * been externalized for browser compatibility"). Replacing the import with a
 * rejected promise removes the specifier from the output entirely; if the
 * branch ever ran, the caller would get a clear error instead of a mystery
 * externalized-module stub.
 *
 * Scope: modern-xlsx's dist chunks only. Our own `await import("node:fs")`
 * in wasm-loader's tryNodeAutoInit MUST stay — it is live code on Node.
 */
const dropNodeFsPromises: EsbuildPlugin = {
  name: "drop-modern-xlsx-node-fs-promises",
  setup(build) {
    build.onLoad(
      { filter: /modern-xlsx[\\/]dist[\\/][^\\/]+\.mjs$/ },
      async (args) => {
        const contents = await fs.promises.readFile(args.path, "utf8");
        if (!contents.includes("node:fs/promises")) return undefined;
        return {
          contents: contents.replace(
            /await import\((['"])node:fs\/promises\1\)/g,
            'await Promise.reject(new Error("node:fs/promises is unavailable: Node-only modern-xlsx APIs (toFile/readFile) are not exported by @marcusok/excel-exporter"))',
          ),
          loader: "js",
        };
      },
    );
  },
};

const shared: Partial<Options> = {
  esbuildPlugins: [rewriteWasmBgUrl, dropNodeFsPromises],
};

// Two configs:
//  - Main entrypoints bundle the engine IN (modern-xlsx + fflate). The package
//    has zero runtime dependencies: consumers install one package, are immune
//    to modern-xlsx's engines.node>=24 declaration, and the wasm binary always
//    ships with the matching JS glue (the exports map re-publishes it).
//  - Worker entrypoint is a SINGLE self-contained file: no imports at all.
//    The default worker URL (new URL("./export.worker.js", import.meta.url) in
//    worker-exporter.ts) is emitted by consumer bundlers as a verbatim asset,
//    and ?url imports copy a single file — either way a chunked worker whose
//    sibling imports are not tracked would 404 in production builds (observed
//    in packages/play/dist: export.worker-*.js referenced a chunk that was
//    never emitted). Splitting stays OFF for this entry; verify with
//    `grep -c "^import" dist/export.worker.js` == 0 after building.
export default defineConfig([
  {
    ...shared,
    entry: {
      index: "src/index.ts",
      "style-presets": "src/style-presets.ts",
      "worker-utils": "src/worker-exporter.ts",
    },
    format: ["esm"],
    dts: { resolve: true },
    splitting: true,
    treeshake: true,
    clean: true,
    sourcemap: true,
    target: "es2022",
    // Browser resolution: without this, tsup defaults to platform "node" and
    // fflate's Node entry bakes a top-level `import { createRequire } from
    // "module"` shim into the bundle, which hard-fails consumer browser
    // builds on Vite 5 ("createRequire is not exported by
    // __vite-browser-external"). We only use fflate's sync APIs, identical in
    // both entries; the Node runtime is unaffected.
    platform: "browser",
  },
  {
    ...shared,
    entry: { "export.worker": "src/workers/export.worker.ts" },
    format: ["esm"],
    dts: false,
    splitting: false,
    treeshake: true,
    sourcemap: true,
    target: "es2022",
    platform: "browser",
    // Everything the worker touches is bundled in: browser module workers
    // cannot resolve bare specifiers (import maps do not apply to
    // WorkerGlobalScope), so the worker script must be self-contained.
    noExternal: ["modern-xlsx", "fflate"],
    clean: false,
  },
]);
