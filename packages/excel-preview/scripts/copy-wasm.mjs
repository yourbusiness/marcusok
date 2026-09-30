/**
 * Post-build step: ship modern-xlsx's WASM binary in this package's dist.
 *
 * The engine layer (xlsx-core, with modern-xlsx bundled in) is bundled INTO
 * this package's dist, so the main-thread loader's default URL resolves here
 * (primary location). The parse worker additionally falls back to the
 * rewritten relative URL inside the bundled worker glue (see
 * @marcusok/xlsx-core/tsup rewriteWasmBgUrl), and self-hosting uses
 * `@marcusok/excel-preview/dist/modern-xlsx.wasm?url`.
 *
 * Runs via the main config's onSuccess hook so watch rebuilds re-copy it too
 * (clean:true wipes dist on watch start — same lesson as excel-exporter 2.1.5).
 */
import { copyFileSync, mkdirSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const distDir = resolve(dirname(fileURLToPath(import.meta.url)), "..", "dist");

// modern-xlsx sits in dependencies (dist/index.d.ts re-exports its types), but
// its runtime code is bundled via xlsx-core — the wasm binary is copied by hand.
const src = resolve(
  dirname(require.resolve("modern-xlsx")),
  "modern-xlsx.wasm",
);

if (!statSync(src, { throwIfNoEntry: false })) {
  throw new Error(
    `[excel-preview] modern-xlsx.wasm not found at ${src}. ` +
      "Run pnpm install first (modern-xlsx is a dependency of this package).",
  );
}

mkdirSync(distDir, { recursive: true });
const dest = resolve(distDir, "modern-xlsx.wasm");
copyFileSync(src, dest);
console.log(`[excel-preview] forwarded wasm -> ${dest}`);
