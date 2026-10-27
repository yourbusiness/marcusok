/**
 * Post-build step: ship pdfjs-dist's runtime assets in this package's dist.
 *
 * Three asset groups are copied verbatim from the pinned pdfjs-dist:
 *  - build/pdf.worker.min.mjs  (self-contained ESM worker, ~1.26MB)
 *  - cmaps/                    (CID-to-Unicode maps, hard requirement for
 *                               CJK PDFs, 169 files)
 *  - standard_fonts/           (Foxit/Liberation fallback fonts for PDFs that
 *                               reference non-embedded standard fonts, 16 files)
 *
 * wasm/ (JPX/JBIG2 image codecs + qcms ICC) and iccs/ are intentionally NOT
 * copied: upstream degrades gracefully when wasmUrl/iccUrl are unset (JS
 * fallback codecs / no color management). Consumers can self-host those and
 * pass the `wasmUrl` / `iccUrl` options.
 *
 * Runs via the main config's onSuccess hook so watch rebuilds re-copy too
 * (clean:true wipes dist on watch start — excel-exporter 2.1.5 lesson).
 */
import { copyFileSync, cpSync, mkdirSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const distDir = resolve(dirname(fileURLToPath(import.meta.url)), "..", "dist");

// pdfjs-dist is a devDependency bundled at build time; locate its install dir
// through Node resolution so pnpm's symlinked layout stays opaque here.
const pkgRoot = dirname(require.resolve("pdfjs-dist/package.json"));

const workerSrc = resolve(pkgRoot, "build", "pdf.worker.min.mjs");
if (!statSync(workerSrc, { throwIfNoEntry: false })) {
  throw new Error(
    `[pdf-preview] pdf.worker.min.mjs not found at ${workerSrc}. ` +
      "Run pnpm install first (pdfjs-dist is a devDependency of this package).",
  );
}

mkdirSync(distDir, { recursive: true });
copyFileSync(workerSrc, resolve(distDir, "pdf.worker.min.mjs"));

for (const dir of ["cmaps", "standard_fonts"]) {
  const src = resolve(pkgRoot, dir);
  if (!statSync(src, { throwIfNoEntry: false })) {
    throw new Error(`[pdf-preview] ${dir}/ not found at ${src}.`);
  }
  // 资产目录内的 LICENSE 文件随目录原样分发（Adobe/Foxit/Liberation 授权
  // 条款要求），不做过滤。
  cpSync(src, resolve(distDir, dir), { recursive: true });
}

console.log(
  `[pdf-preview] forwarded worker + cmaps/ + standard_fonts/ -> ${distDir}`,
);
