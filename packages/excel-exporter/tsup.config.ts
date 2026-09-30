import { defineConfig, type Options } from "tsup";
// 引擎 esbuild 插件自共享底层包复用（wasm-glue URL 重写 / node:fs/promises
// 剔除）。xlsx-core 已转为仓库私有包（不再发布 npm），仅作本包的构建期
// 依赖：主入口与 worker 都把它的 dist 整体打进本包产物。core 需先构建
// （turbo ^build 已编排）。
import { rewriteWasmBgUrl, dropNodeFsPromises } from "@marcusok/xlsx-core/tsup";

const shared: Partial<Options> = {
  esbuildPlugins: [rewriteWasmBgUrl, dropNodeFsPromises],
};

// Two configs:
//  - Main entrypoints bundle the shared layers (@marcusok/xlsx-core, the
//    modern-xlsx engine layer; @marcusok/progress-overlay, the overlay UI)
//    INTO dist. Both are private workspace packages, so consumers depend on
//    this package alone — the npm release surface is the two business
//    packages, with no cross-package version coupling. Accepted trade-off
//    (documented in the docs-site package-relationships guide): a page using
//    both business packages loads two engine instances and two WASM binaries
//    on the main thread; most admin pages use only one of the two, and
//    bundler content-hash assets plus HTTP cache usually dedupe the binary
//    transfer. Workers always carried their own copy anyway (see below).
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
      // 遮罩是独立入口：index.ts 不引用它，因此不用该能力的调用方打包体积
      // 不变（仅多一个可按需 import 的子路径文件）。progress-overlay 的实现
      // 与类型都内联进本入口（它是私有包，消费方无法单独安装）。
      overlay: "src/overlay.ts",
    },
    format: ["esm"],
    dts: { resolve: true },
    splitting: true,
    treeshake: true,
    clean: true,
    sourcemap: true,
    target: "es2022",
    // 引擎层（xlsx-core，含 modern-xlsx）与遮罩层（progress-overlay）打进
    // 本包产物——二者已私有化，发布面只剩两个业务包。dts 侧实测：本包自身
    // 的类型声明只"引用"引擎类型（不作纯 re-export），rollup-dts 能顺着
    // core 的 d.ts 解析到 modern-xlsx 并整体内联（dist 无 'modern-xlsx'
    // 外部引用）；但内联行为不保证稳定（excel-preview 的纯 re-export 面就
    // 保留外部引用），package.json 的 dependencies 仍保留 modern-xlsx
    // （仅类型用、与打包版本精确钉同版）兜底消费方 TS 解析。
    noExternal: ["@marcusok/xlsx-core", "@marcusok/progress-overlay"],
    // copy-wasm 挂在 onSuccess 而非只在 "build" script：本配置 clean:true，
    // watch 模式首次构建即清空 dist（连带删掉上次 build 复制的 wasm）且
    // 此后不再回补，开着 pnpm dev 时 Node 自动初始化/集成测试会因
    // dist/modern-xlsx.wasm 缺失而失败。onSuccess 在普通构建与 watch 的
    // 每次重建后都执行，"build" 与 "dev" 两条链路由此统一回补。
    // 引擎打进本包后，这份二进制就是主线程 loader 默认 URL 的正主（不再是
    // deprecated 转发——@marcusok/xlsx-core 已私有，外部没有另一个路径）。
    onSuccess: "node scripts/copy-wasm.mjs",
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
    // WorkerGlobalScope), so the worker script must be self-contained. core
    // carries modern-xlsx in its dist (bundle-IN), so this one entry suffices.
    noExternal: ["@marcusok/xlsx-core", "fflate"],
    clean: false,
  },
]);
