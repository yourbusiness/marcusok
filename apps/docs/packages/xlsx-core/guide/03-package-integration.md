# Integrating the Core into a Package

This page is for authors of packages **built on top of** `@marcusok/xlsx-core` — the two esbuild plugins the core exposes for `tsup`, how to wire them, how the WASM asset is located, and the `engines` constraints that come with depending on the core. For application-side usage stay on the [loader guide](/packages/xlsx-core/guide/01-loader).

## Why a package needs build plugins at all

The engine ships as JS glue plus a `.wasm` binary, and both bring assumptions that break in a consumer's browser build:

- The wasm-bindgen glue resolves its binary with `new URL('modern_xlsx_wasm_bg.wasm', import.meta.url)` — a filename **no** @marcusok package ships (the binary is forwarded as `dist/modern-xlsx.wasm`). The code path is dead at runtime because the loader always passes an explicit URL, but bundlers statically analyse every literal `new URL(..., import.meta.url)` and warn when the file does not exist (Vite: _"doesn't exist at build time, it will remain unchanged…"_).
- modern-xlsx's Node-only file APIs (`toFile`, `readFile`) dynamic-import `node:fs/promises`. @marcusok packages export buffer APIs only (`toBuffer` / `readBuffer`), so those imports are dead branches — but they survive as literal dynamic imports, and consumer browser builds warn about them (Vite 5 / VitePress: _"Module fs/promises has been externalized for browser compatibility"_).

`@marcusok/xlsx-core/tsup` exports one plugin per problem, and both are applied at **build time of your package**, so the bad code never reaches the consumer's bundler.

| Plugin               | What it rewrites                                                                                                                                   | Why it is safe                                                                                                                                                                                                                                                                                                                          |
| -------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `rewriteWasmBgUrl`   | `new URL('modern_xlsx_wasm_bg.wasm', import.meta.url)` → `new URL(["..", "dist", "modern-xlsx.wasm"].join("/"), import.meta.url)` in the glue file | The replacement is deliberately **not** a string literal: neither esbuild nor a downstream bundler matches computed specifiers, so nothing is asset-analysed or emitted twice. If the branch ever ran it would still resolve — from the glue file and from your bundling package's `dist`, `../dist/modern-xlsx.wasm` is a shipped copy |
| `dropNodeFsPromises` | `await import('node:fs/promises')` → `await Promise.reject(new Error(...))` in modern-xlsx's own dist chunks                                       | The engine's Node-only APIs are not exported by @marcusok packages; a dead branch that fails loudly beats a mystery externalized-module stub                                                                                                                                                                                            |

Scope matters for the second one: the filter is modern-xlsx's dist chunks only. Your own live `await import(...)` of a Node built-in must stay — the core's Node auto-init uses exactly that, with a computed specifier so browser bundlers never see a static `node:fs` import.

## Wiring them into `tsup.config.ts`

Both consuming packages use the same shape — a shared options object spread into every config, so the worker entry gets the plugins too:

```ts
import { defineConfig, type Options } from "tsup";
import { rewriteWasmBgUrl, dropNodeFsPromises } from "@marcusok/xlsx-core/tsup";

const shared: Partial<Options> = {
  esbuildPlugins: [rewriteWasmBgUrl, dropNodeFsPromises],
};

export default defineConfig([
  {
    ...shared,
    entry: { index: "src/index.ts" },
    format: ["esm"],
    dts: { resolve: true },
    // Keep the core external: a page using several @marcusok packages then
    // loads one engine instance and one WASM binary on the main thread.
    external: ["@marcusok/xlsx-core"],
    platform: "browser",
    // Re-copy the binary after every build AND every watch rebuild (see below).
    onSuccess: "node scripts/copy-wasm.mjs",
  },
  {
    ...shared,
    entry: { "my.worker": "src/workers/my.worker.ts" },
    format: ["esm"],
    dts: false,
    splitting: false, // single self-contained file — see below
    platform: "browser",
    // Workers bundle the core IN: browser module workers cannot resolve bare
    // specifiers, so the script must be self-contained.
    noExternal: ["@marcusok/xlsx-core"],
    clean: false,
  },
]);
```

`@marcusok/xlsx-core/tsup` is a **Node-only** entry (it imports `node:fs` to read the glue file at build time), so it belongs in your `devDependencies`-resolved tsup config and nowhere else. Because the core bundles the engine in, you do **not** need `modern-xlsx` as a dependency, and you do not need a `noExternal: ["modern-xlsx"]` entry either — `noExternal: ["@marcusok/xlsx-core"]` brings the whole engine along.

**External on the main entry, bundled in the worker.** That split is the whole point of the core: the main thread shares one engine instance (and one WASM binary) across every @marcusok package on the page, while a worker script must be a single self-contained file. Verify the worker with:

```bash
grep -c "^import" dist/my.worker.js   # must be 0
```

A chunked worker whose sibling imports are not tracked will 404 in production builds, which is why `splitting` stays off for worker entries.

## Locating and publishing the WASM asset

The loader's default URL is `new URL("./modern-xlsx.wasm", import.meta.url)` — the binary has to sit next to your package's entry after your build. Forward it with a post-build copy step, for the same reason the core does:

```js
// scripts/copy-wasm.mjs
import { copyFileSync, mkdirSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const distDir = resolve(dirname(fileURLToPath(import.meta.url)), "..", "dist");
const src = resolve(
  dirname(require.resolve("modern-xlsx")),
  "modern-xlsx.wasm",
);

if (!statSync(src, { throwIfNoEntry: false })) {
  throw new Error(
    `modern-xlsx.wasm not found at ${src}. Run pnpm install first.`,
  );
}

mkdirSync(distDir, { recursive: true });
copyFileSync(src, resolve(distDir, "modern-xlsx.wasm"));
```

`require.resolve("modern-xlsx")` resolves the package entry, whose directory holds the binary — the same resolution the loader's Node auto-init relies on. Note what that implies: both consuming packages declare `modern-xlsx` as a **devDependency** of their own purely so this script can find the source binary (they still resolve `@marcusok/xlsx-core` at runtime, never `modern-xlsx`). Under pnpm's strict `node_modules` layout a dependency of the core is not resolvable from your package, so the script cannot borrow the core's copy.

Two more details that were learned the hard way in this repo:

- Hook it to tsup's `onSuccess`, not only to the `build` npm script. The main config uses `clean: true`, so a `--watch` run wipes `dist/` (including the previously copied wasm) at startup and would never put it back; `onSuccess` runs after every rebuild in both `build` and `dev`.
- Publish it under your package's `exports` map (a plain file entry such as `"./dist/modern-xlsx.wasm": "./dist/modern-xlsx.wasm"`). That is what makes `your-package/dist/modern-xlsx.wasm?url` resolvable in Vite / webpack 5 — the engine's own exports map deliberately omits its wasm subpaths, so a deep import into `modern-xlsx` cannot work.

`new URL(<file>, import.meta.url)` is then the only locator you need: bundlers rewrite the expression and emit the file as a hashed asset, and Node resolves it through `node_modules`. Reading it from disk in Node (which is what the loader's auto-init does) additionally requires the package's `dist/` to be on disk — a server build that inlines dependencies without emitting the asset has to keep the package external, pass `configureWasm({ wasmUrl })`, or copy the file next to the bundle.

## `engines` and the pinned dependency

- The core declares `engines.node >= 22`. The engine it bundles declares `>= 24`, but that range is inert at runtime — nothing imports `modern-xlsx` — which is precisely why the core redeclares a range its consumers can satisfy.
- `modern-xlsx` nonetheless stays declared as an exact-pinned dependency of the core, for **type resolution only**: tsup's dts pass cannot inline the re-exported types (modern-xlsx is an exports-map-only package with no top-level `types` entry, which tsup's dts resolver cannot resolve), so the published `.d.ts` keeps `from "modern-xlsx"` imports and TypeScript needs the declaration to resolve them. Do not "clean it up" as unused.
- Consequence for installers: with `engine-strict=true` enabled on Node < 24 an install is rejected (npm / pnpm default is a warning only). The [package README](https://www.npmjs.com/package/@marcusok/xlsx-core) documents the full rationale.

## Handling engine errors

If your package surfaces engine failures to callers, branch on `ModernXlsxError`'s `code` rather than on message text — messages change with the input shape, codes are the stable contract. Map the codes you care about onto your own error vocabulary, exactly as `@marcusok/excel-preview` does in its [error normalizer](/packages/excel-preview/api/01-create-preview#previewerror-codes). The four value-exported constants ([typed errors](/packages/xlsx-core/guide/02-engine-surface#typed-errors)) cover the cases that matter most; anything else has to be matched by its literal string.
