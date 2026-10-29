import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { existsSync, readFileSync, readdirSync, realpathSync } from "node:fs";
import { createRequire } from "node:module";
import {
  isFile,
  matchExportTarget,
  readExportsMap,
  srcCandidatesFromDistTarget,
  type SourcePackage,
} from "./src/vite/workspace-resolver.ts";
import { rewriteAssetLiterals } from "./src/vite/src-asset-overrides.ts";

const rootDir = dirname(fileURLToPath(import.meta.url));
// play 是应用，位于 apps/ 下；workspace 业务包统一在仓库根的 packages/
// （上两级到仓库根再进 packages）。
const packagesDir = resolve(rootDir, "..", "..", "packages");

/**
 * Escape regex metacharacters in a package name before splicing it into a
 * RegExp. Names like "pdf.js" would otherwise silently mis-anchor the alias
 * patterns (the dot matches any character) instead of failing loudly.
 */
function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Source entry of a workspace package. The monorepo convention is
 * src/index.ts (fallback: src/index.tsx); packages that don't follow it get
 * a startup warning below and resolve through their dist instead (no HMR).
 */
function sourceEntry(dir: string): string {
  const ts = resolve(dir, "src/index.ts");
  return isFile(ts) ? ts : resolve(dir, "src/index.tsx");
}

/**
 * Workspace packages that expose a src/index.ts(.tsx) source entry. These get
 * HMR-friendly aliases; everything else is handled by normal node/vite
 * resolution against the workspace symlink (i.e. built dist).
 */
// （play 自身在 apps/ 下，不在 packagesDir 扫描范围内，无需排除。）
const sourcePackages: SourcePackage[] = readdirSync(packagesDir)
  .filter(
    (name) =>
      isFile(resolve(packagesDir, name, "src/index.ts")) ||
      isFile(resolve(packagesDir, name, "src/index.tsx")),
  )
  .map((name) => ({
    name,
    dir: resolve(packagesDir, name),
    exportsMap: readExportsMap(resolve(packagesDir, name)),
  }));

// Fail loud instead of silently losing source HMR: warn for every declared
// @marcusok/* dependency that doesn't follow the src/index.ts convention.
const declaredDeps: string[] = (() => {
  try {
    const json = JSON.parse(
      readFileSync(resolve(rootDir, "package.json"), "utf8"),
    ) as { dependencies?: Record<string, string> };
    return Object.keys(json.dependencies ?? {})
      .filter((name) => name.startsWith("@marcusok/"))
      .map((name) => name.replace(/^@marcusok\//, ""));
  } catch {
    return [];
  }
})();

for (const name of declaredDeps) {
  if (!sourcePackages.some((pkg) => pkg.name === name)) {
    console.warn(
      `[play] "@marcusok/${name}" 未检测到 src/index.ts(.tsx)，主入口将走 dist 解析（无源码 HMR）。请按约定提供 src/index.ts。`,
    );
  }
}

/**
 * Third-party packages whose "exports" map doesn't expose certain runtime
 * subpaths (e.g. ./wasm/*). When a demo imports `<pkg>/<subpath>?url` and
 * Vite's export-map resolution would reject it, we rewrite to the physical
 * file on disk if it exists.
 *
 * To support a new third-party subpath, add an entry here; resolver logic
 * doesn't need to change. 现状：唯一条目是 @silurus/ooxml（ppt-preview 的
 * 引擎；主要价值在 excludeFromOptimizeDeps + fs.allow 放行，详见条目注释，
 * 子路径改写只是顺带兜底）。@marcusok/excel-exporter >= 2.0 不需要条目：
 * 它把运行资产显式再发布进自身 exports map（`./dist/export.worker.js`、
 * `./dist/modern-xlsx.wasm`，无 `./dist/*` 通配），给它新增 dist 资产时应
 * 先补 exports，而不是加 override。
 */
const externalOverrides: {
  pkg: string;
  dir: string;
  excludeFromOptimizeDeps?: boolean;
}[] = [
  {
    // ppt-preview 的渲染引擎：chunk 内 `new URL("pptx_parser_bg.wasm",
    // import.meta.url)` 依赖模块真实文件位置定位 wasm。optimizeDeps 预打包
    // 会把 chunk 搬进 .vite/deps，相对解析随即指向不存在的路径（ppt-preview
    // S1 spike 的 play dev 冒烟结论）——排除预打包，让 chunk 与 wasm 留在
    // 上游 dist 内按真实相对路径加载。dir 的 fs.allow 同步放行（见 server
    // 配置）：排除预打包后 dev 直接 serve node_modules 物理文件，默认 allow
    // 不含 .pnpm 深层路径。
    pkg: "@silurus/ooxml",
    dir: (() => {
      try {
        // @silurus/ooxml 是 ppt-preview 的 devDependency，play 自身没有声明
        // ——经 ppt-preview 的依赖树解析（exports 含 /pptx 子路径）；realpath
        // 穿透 pnpm 符号链接拿 .pnpm 物理目录
        const req = createRequire(
          resolve(packagesDir, "ppt-preview", "package.json"),
        );
        return dirname(realpathSync(req.resolve("@silurus/ooxml/pptx")));
      } catch {
        return "";
      }
    })(),
    excludeFromOptimizeDeps: true,
  },
];

// Fail loud, symmetric with the @marcusok/* check above: an entry whose dir
// cannot be resolved would otherwise be skipped silently and its subpath
// imports would 500 at dev time with no hint why.
for (const { pkg, dir } of externalOverrides) {
  if (!dir) {
    console.warn(
      `[play] externalOverrides 声明了 "${pkg}"，但无法从 ${rootDir} 解析到该包目录（检查它是否已安装、exports 是否仅含 import 条件）。相关子路径导入将走正常解析并可能 404/500。`,
    );
  }
}

// Main-entry aliases: @marcusok/<pkg> -> src/index.ts(.tsx) (HMR-friendly source).
// Anchored regex so it does NOT swallow subpaths like /dist/export.worker.js.
const mainAliases = sourcePackages.map(({ name, dir }) => ({
  find: new RegExp(`^@marcusok/${escapeRegExp(name)}$`),
  replacement: sourceEntry(dir),
}));

// src-alias asset mapping: @marcusok/excel-exporter locates its wasm/worker
// via `new URL(<literal>, import.meta.url)`, relative to the published dist
// entry — the assets sit in dist/, next to the entry. The aliases above point
// the main entry at src/, so those literals resolve into src/ where the files
// do not exist. A transform rewrites each literal into an injected
// `<pkg>/dist/<file>?url` import — dev resolves it through marcusok-resolver
// to the on-disk package file (served same-origin), build emits it via the
// standard ?url asset pipeline (see src-asset-overrides.ts for why the
// previous `../dist/<file>` relative rewrite was replaced).
const srcAssetOverrides: {
  pkg: string;
  dir: string;
  files: string[];
}[] = [
  {
    pkg: "@marcusok/excel-exporter",
    dir: resolve(packagesDir, "excel-exporter"),
    files: ["export.worker.js"],
  },
  // wasm 字面量随 WasmLoader 迁入 core（excel-preview 与 exporter 共用）；
  // 预览包另有自包含 parse worker。
  {
    pkg: "@marcusok/xlsx-core",
    dir: resolve(packagesDir, "xlsx-core"),
    files: ["modern-xlsx.wasm"],
  },
  {
    pkg: "@marcusok/excel-preview",
    dir: resolve(packagesDir, "excel-preview"),
    files: ["parse.worker.js"],
  },
  // pdf-preview 以 worker 单文件字面量为唯一锚点（assets.ts）；cmaps 与
  // standard_fonts 从锚点 URL 推导同层目录（dev 下解析到包内 dist 磁盘文件，
  // 同源可 fetch——rewriteAssetLiterals 不支持目录，目录不经字面量改写）。
  {
    pkg: "@marcusok/pdf-preview",
    dir: resolve(packagesDir, "pdf-preview"),
    files: ["pdf.worker.min.mjs"],
  },
  // ppt-preview 的 assets.ts 以 wasm 为唯一静态锚点；兜底分支以 anchor
  // 变量为 base（源码里无 ../dist/ 静态字面量，不会漏改写——包内注释有
  // 说明）。上游 @silurus/ooxml chunk 内的 wasm 字面量不经此处：那部分由
  // externalOverrides 的 excludeFromOptimizeDeps 保真实路径。
  {
    pkg: "@marcusok/ppt-preview",
    dir: resolve(packagesDir, "ppt-preview"),
    files: ["pptx_parser_bg.wasm"],
  },
];

const srcAssetOverridePlugin: Plugin = {
  name: "marcusok-src-asset-overrides",
  enforce: "pre",
  transform(code, id) {
    // Vite module ids use forward slashes even on Windows; normalize both
    // sides before comparing, or the startsWith check silently never matches.
    const norm = (p: string) => p.replaceAll("\\", "/");
    const override = srcAssetOverrides.find(({ dir }) =>
      norm(id).startsWith(`${norm(resolve(dir, "src"))}/`),
    );
    if (!override) return null;
    const rewritten = rewriteAssetLiterals(code, override.pkg, override.files);
    return rewritten === null ? null : { code: rewritten, map: null };
  },
};

export default defineConfig({
  resolve: { alias: mainAliases },
  plugins: [
    react(),
    srcAssetOverridePlugin,
    {
      name: "marcusok-resolver",
      enforce: "pre",
      resolveId(source) {
        const queryIdx = source.indexOf("?");
        const clean = queryIdx === -1 ? source : source.slice(0, queryIdx);
        const query = queryIdx === -1 ? "" : source.slice(queryIdx);

        // @marcusok/<pkg>/<sub...> resolution:
        //  1. "dist/..." requests are build artifacts (workers, wasm) ->
        //     resolve straight to the on-disk file; never source.
        //  2. Flat source file: src/<sub>.ts(/.tsx) or src/<sub>/index.*.
        //  3. Exports-map-aware: map "./<sub>" (or a "*" wildcard key) to
        //     its dist target, then back to the matching src file when it
        //     exists. Keeps subpaths like "./styles" -> style-presets.ts
        //     HMR-friendly even when the export key and source file differ.
        //  4. Bare dist artifact as a safety net for subpaths the upstream
        //     package omits from its "exports" map.
        //  5. Otherwise return null and let normal resolution handle it.
        //  ?url is preserved everywhere so Vite serves files as URLs rather
        //  than importing them as modules.
        const sub = clean.match(/^@marcusok\/([^/]+)\/(.+)$/);
        if (sub) {
          const pkgName = sub[1]!;
          const subpath = sub[2]!;
          const pkg = sourcePackages.find((p) => p.name === pkgName);
          if (pkg) {
            if (subpath.startsWith("dist/")) {
              const distCandidate = resolve(pkg.dir, subpath);
              if (isFile(distCandidate)) return distCandidate + query;
              return null;
            }
            const srcFlat = [
              `${subpath}.ts`,
              `${subpath}.tsx`,
              `${subpath}/index.ts`,
              `${subpath}/index.tsx`,
            ].map((rel) => resolve(pkg.dir, "src", rel));
            for (const candidate of srcFlat) {
              if (isFile(candidate)) return candidate + query;
            }
            const exportTarget = matchExportTarget(pkg.exportsMap, subpath);
            if (exportTarget) {
              for (const rel of srcCandidatesFromDistTarget(exportTarget)) {
                const candidate = resolve(pkg.dir, "src", rel);
                if (isFile(candidate)) return candidate + query;
              }
            }
            const distCandidate = resolve(pkg.dir, "dist", subpath);
            if (isFile(distCandidate)) return distCandidate + query;
          }
          return null;
        }

        // Third-party packages with unexported subpaths: rewrite to the
        // physical file on disk if it exists. Safe because we only rewrite
        // when the file actually exists at the subpath relative to the
        // package root; subpaths covered by "exports" (e.g. "./lite") live
        // under dist/ and won't match existsSync, so they fall through to
        // normal resolution.
        for (const { pkg, dir } of externalOverrides) {
          if (!dir) continue;
          const m = clean.match(new RegExp(`^${escapeRegExp(pkg)}/(.+)$`));
          if (m) {
            const physical = resolve(dir, m[1]!);
            if (existsSync(physical)) return physical + query;
          }
        }
        return null;
      },
    },
  ],
  optimizeDeps: {
    // Packages that use new URL(..., import.meta.url) for wasm or workers
    // break under pre-bundling; exclude all declared entries that need it.
    exclude: externalOverrides
      .filter((e) => e.excludeFromOptimizeDeps)
      .map((e) => e.pkg),
  },
  server: {
    port: 5173,
    // Fail loudly instead of silently hopping to 5174/5175 when the port is
    // already taken (multiple dev instances are easy to mix up otherwise).
    // Override with `vite --port <n>` if 5173 is genuinely occupied.
    strictPort: true,
    fs: {
      // Serve the play itself (incl. its node_modules) plus the
      // workspace packages that follow the source convention. New packages
      // are picked up automatically once they add src/index.ts(.tsx).
      // externalOverrides 条目（如 @silurus/ooxml）排除预打包后按物理路径
      // serve node_modules 深层文件，需一并放行。
      allow: [
        rootDir,
        ...sourcePackages.map((pkg) => pkg.dir),
        ...externalOverrides.map((e) => e.dir).filter(Boolean),
      ],
    },
  },
});
