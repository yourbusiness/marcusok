/**
 * runtimeAssets 拷贝机制（从 config.ts 抽出以便测试）：
 * 把注册表声明的 dist 资产（wasm/worker 单文件、cmaps/standard_fonts
 * 整目录）拷入 public/，新包只需在 registry 声明 runtimeAssets、无需改
 * 构建配置。
 *
 * 两种形态：
 * - file：单文件，`file` 相对 resolveFrom 解析出的**入口所在目录**
 *   （如 wasm 与包入口同层）；
 * - dir：整目录递归拷贝，`dir` 相对 resolveFrom 的**包根目录**解析
 *   （pdfjs-dist 的 cmaps/ 在包根、不在 build/ 下）。
 *
 * 目录条目**不得**进 stableAssetNames（见 stableAssetNamesFor）：该集合
 * 只对"经 Rollup emit 的文件"有意义（让 emit 不带 hash、与 public 拷贝
 * 路径对齐），cmaps 走运行时 fetch、不经 emit，进集合只会产出无意义的
 * basename 并可能误命中同名 emit 资产。
 */
import { copyFileSync, cpSync, mkdirSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import { basename, dirname, join } from "node:path";
import type { PackageEntry, RuntimeAsset } from "./registry";

const require = createRequire(import.meta.url);

/**
 * Resolve the directory containing a package's dist assets. When `through`
 * is set, the resolution happens within that package's dependency context
 * (e.g. pdfjs-dist is a devDependency of @marcusok/pdf-preview, not of the
 * docs app itself).
 */
function resolveAssetDir(asset: RuntimeAsset): string {
  if (asset.through) {
    const throughReq = createRequire(
      require.resolve(join(asset.through, "package.json")),
    );
    return dirname(
      throughReq.resolve(
        "file" in asset
          ? asset.resolveFrom
          : join(asset.resolveFrom, "package.json"),
      ),
    );
  }
  return dirname(
    require.resolve(
      "file" in asset
        ? asset.resolveFrom
        : join(asset.resolveFrom, "package.json"),
    ),
  );
}

/**
 * Copy every package's declared runtimeAssets from their dist directories
 * into publicDir.
 */
export function copyRuntimeAssets(
  pkgs: readonly PackageEntry[],
  publicDir: string,
): void {
  for (const p of pkgs) {
    for (const asset of p.runtimeAssets ?? []) {
      const src = join(
        resolveAssetDir(asset),
        "file" in asset ? asset.file : asset.dir,
      );
      const dest = join(publicDir, asset.to);
      if (!statSync(src, { throwIfNoEntry: false })) {
        throw new Error(
          `[docs] runtime asset not found: ${asset.resolveFrom}${asset.through ? ` (through ${asset.through})` : ""}/${"file" in asset ? asset.file : asset.dir} ` +
            `(declared by ${p.npmName} runtimeAssets). ` +
            "Run `pnpm build:docs` from the repo root (turbo builds deps first).",
        );
      }
      mkdirSync(dirname(dest), { recursive: true });
      if ("file" in asset) {
        copyFileSync(src, dest);
      } else {
        // 目录原样递归拷贝（含资产自带的 LICENSE 文件——授权条款随分发）
        cpSync(src, dest, { recursive: true });
      }
    }
  }
}

/**
 * Destination basenames of file-form runtime assets (dir-form entries are
 * excluded by design — see the module header). Rollup emits these without a
 * content hash so they land at the same path as the copies in public/.
 */
export function stableAssetNamesFor(
  pkgs: readonly PackageEntry[],
): Set<string> {
  return new Set(
    pkgs
      .flatMap((p) => p.runtimeAssets ?? [])
      .filter((a): a is Extract<RuntimeAsset, { file: string }> => "file" in a)
      .map((a) => basename(a.to)),
  );
}
