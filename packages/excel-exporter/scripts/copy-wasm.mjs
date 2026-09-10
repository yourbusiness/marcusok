/**
 * Post-build step: forward modern-xlsx's WASM binary into this package's dist.
 *
 * modern-xlsx's "exports" map omits its wasm subpaths, so consumer bundlers
 * (Vite, webpack) reject deep imports like `modern-xlsx/dist/modern-xlsx.wasm?url`.
 * Re-publishing the binary under our own exports map makes
 * `@marcusok/excel-exporter/dist/modern-xlsx.wasm?url` resolvable, which is the
 * zero-plugin asset path documented in the README.
 *
 * Runs after `tsup` (see the "build" script); dist/modern-xlsx.wasm ships with
 * the package via the "files": ["dist"] entry.
 */
import { copyFileSync, mkdirSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const distDir = resolve(dirname(fileURLToPath(import.meta.url)), "..", "dist");

// Same resolution the Node auto-init in src/wasm-loader.ts relies on: resolving
// the package entry yields its dist directory, which holds the wasm binary.
const src = resolve(
  dirname(require.resolve("modern-xlsx")),
  "modern-xlsx.wasm",
);

if (!statSync(src, { throwIfNoEntry: false })) {
  throw new Error(
    `[excel-exporter] modern-xlsx.wasm not found at ${src}. ` +
      "Run pnpm install first (modern-xlsx is a devDependency of this package).",
  );
}

mkdirSync(distDir, { recursive: true });
const dest = resolve(distDir, "modern-xlsx.wasm");
copyFileSync(src, dest);
console.log(`[excel-exporter] forwarded wasm -> ${dest}`);
