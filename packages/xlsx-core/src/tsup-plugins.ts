import fs from "node:fs";
import type { Options } from "tsup";

/** esbuild plugin type, derived from tsup's own option shape (esbuild itself is a transitive dep under pnpm). */
export type EsbuildPlugin = NonNullable<Options["esbuildPlugins"]>[number];

/**
 * modern-xlsx's wasm-bindgen glue (wasm/modern_xlsx_wasm.js) defaults to
 * `new URL('modern_xlsx_wasm_bg.wasm', import.meta.url)` — a filename no
 * @marcusok package ships (the binary is forwarded as dist/modern-xlsx.wasm;
 * sha256-verified equal). Our code always passes an explicit URL, so the
 * branch is dead at runtime, but consumer bundlers statically analyze every
 * `new URL(<literal>, import.meta.url)` and warn when the file is missing
 * (Vite: "doesn't exist at build time, it will remain unchanged...").
 *
 * The replacement is deliberately NOT a string literal (array join): neither
 * esbuild nor downstream bundlers match computed specifiers, so nothing is
 * asset-analyzed or emitted twice. If the branch ever executes at runtime it
 * still resolves correctly — ["..","dist","modern-xlsx.wasm"].join("/") from
 * the glue file and from the bundling package's dist both point at a shipped
 * copy (each @marcusok package re-forwards the binary into its own dist).
 */
export const rewriteWasmBgUrl: EsbuildPlugin = {
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
 * `node:fs/promises`. @marcusok packages never export those APIs (only
 * buffer APIs: toBuffer/readBuffer), so in the bundled output the imports
 * are dead branches — but they survive as literal dynamic imports, and
 * consumer browser builds warn about them (Vite 5/VitePress: "Module
 * fs/promises has been externalized for browser compatibility"). Replacing
 * the import with a rejected promise removes the specifier from the output
 * entirely; if the branch ever ran, the caller would get a clear error
 * instead of a mystery externalized-module stub.
 *
 * Scope: modern-xlsx's dist chunks only. A package's own `await
 * import("node:fs")` in wasm-loader's tryNodeAutoInit MUST stay — it is live
 * code on Node (kept alive there via a computed specifier).
 */
export const dropNodeFsPromises: EsbuildPlugin = {
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
            'await Promise.reject(new Error("node:fs/promises is unavailable: Node-only modern-xlsx APIs (toFile/readFile) are not exported by @marcusok packages"))',
          ),
          loader: "js",
        };
      },
    );
  },
};
