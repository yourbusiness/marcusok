import { defineConfig, type Options } from "tsup";
// 引擎 esbuild 插件自共享底层包复用（wasm-glue URL 重写 / node:fs/promises
// 剔除），不再在本包内各留一份拷贝。core 需先构建（turbo ^build 已编排）。
import { rewriteWasmBgUrl, dropNodeFsPromises } from "@marcusok/xlsx-core/tsup";

const shared: Partial<Options> = {
  esbuildPlugins: [rewriteWasmBgUrl, dropNodeFsPromises],
};

// Two configs:
//  - Main entrypoints depend on @marcusok/xlsx-core (the shared modern-xlsx
//    engine layer, bundled-IN there) and keep it EXTERNAL: consumers install
//    core alongside this package, and a page using several @marcusok
//    packages loads one engine instance + one WASM binary on the main
//    thread. Everything else this package needs is bundled in (fflate for
//    the fast stream), keeping the runtime dependency set at exactly one
//    same-scope package.
//  - Worker entrypoint is a SINGLE self-contained file: no imports at all.
//    The default worker URL (new URL("./export.worker.js", import.meta.url) in
//    worker-exporter.ts) is emitted by consumer bundlers as a verbatim asset,
//    and ?url imports copy a single file — either way a chunked worker whose
//    sibling imports are not tracked would 404 in production builds (observed
//    in packages/play/dist: export.worker-*.js referenced a chunk that was
//    never emitted). Splitting stays OFF for this entry; verify with
//    `grep -c "^import" dist/export.worker.js` == 0 after building.
//    modern-xlsx itself no longer appears here directly: it is inlined in
//    core's dist, so bundling core brings the whole engine along.
export default defineConfig([
  {
    ...shared,
    entry: {
      index: "src/index.ts",
      "style-presets": "src/style-presets.ts",
      "worker-utils": "src/worker-exporter.ts",
      // 遮罩是独立入口：index.ts 不引用它，因此不用该能力的调用方打包体积
      // 不变（仅多一个可按需 import 的子路径文件）。
      overlay: "src/overlay.ts",
    },
    format: ["esm"],
    dts: { resolve: true },
    splitting: true,
    treeshake: true,
    clean: true,
    sourcemap: true,
    target: "es2022",
    // 主入口不打包 core：共享单份引擎与 wasm 的关键（见文件头注释）。
    external: ["@marcusok/xlsx-core"],
    // copy-wasm 挂在 onSuccess 而非只在 "build" script：本配置 clean:true，
    // watch 模式首次构建即清空 dist（连带删掉上次 build 复制的 wasm）且
    // 此后不再回补，开着 pnpm dev 时 Node 自动初始化/集成测试会因
    // dist/modern-xlsx.wasm 缺失而失败。onSuccess 在普通构建与 watch 的
    // 每次重建后都执行，"build" 与 "dev" 两条链路由此统一回补。
    // 注：这是对既有消费方的 deprecated 转发（新路径为
    // @marcusok/xlsx-core/dist/modern-xlsx.wasm），README 已标注。
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
