/**
 * Post-build step: ship modern-xlsx's WASM binary in this package's dist.
 *
 * 本包已私有化（不再发布 npm），dist 里这份二进制的消费者都在仓内：业务包
 * tsup 把本包 dist 的 JS 胶水打进各自产物后，由各自的 copy-wasm 把二进制
 * 转发进业务包 dist（对外唯一的分发面）；play 的 src 联调链路
 * （srcAssetOverrides 重写）与本包的 Node 集成测试（wasm-loader 默认 URL）
 * 直接读本目录下这份。曾经的 `@marcusok/xlsx-core/dist/modern-xlsx.wasm`
 * 对外自托管路径已随私有化废弃（exports 条目一并移除）。
 *
 * Runs after `tsup` via the main config's onSuccess hook (see tsup.config.ts),
 * covering both the one-shot "build" script and every "dev" watch rebuild —
 * otherwise the clean:true wipe at watch startup leaves dist without the wasm
 * (Node auto-init / integration tests then fail). dist/modern-xlsx.wasm ships
 * with the package via the "files": ["dist"] entry.
 *
 * (Migrated verbatim from excel-exporter — this package is now the engine
 * integration point, so the forwarding lives here.)
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
    `[xlsx-core] modern-xlsx.wasm not found at ${src}. ` +
      "Run pnpm install first (the engine's runtime is bundled from this copy at build time).",
  );
}

mkdirSync(distDir, { recursive: true });
const dest = resolve(distDir, "modern-xlsx.wasm");
copyFileSync(src, dest);
console.log(`[xlsx-core] forwarded wasm -> ${dest}`);
