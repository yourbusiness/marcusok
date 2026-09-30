import { defineConfig, type Options } from "tsup";
import { rewriteWasmBgUrl, dropNodeFsPromises } from "./src/tsup-plugins";

// Two configs:
//  - Main entry bundles the engine IN (modern-xlsx). This package is a
//    private workspace layer (no longer published): the business packages
//    bundle its dist into their own builds (noExternal), so their consumers
//    are immune to modern-xlsx's engines.node>=24 declaration and the wasm
//    binary always ships with the matching JS glue next to the bundled
//    entry. Within the repo, this dist is also the bundling source for the
//    business packages' self-contained workers.
//  - The ./tsup subpath exposes the esbuild plugins the business packages'
//    tsup configs reuse (build-time only — it never ships to npm consumers
//    any more, but the workspace devDependency keeps resolving it). It must
//    run on Node (imports node:fs), hence its own platform.
export default defineConfig([
  {
    esbuildPlugins: [rewriteWasmBgUrl, dropNodeFsPromises],
    entry: { index: "src/index.ts" },
    format: ["esm"],
    dts: { resolve: true },
    // noExternal 必不可少：modern-xlsx 声明在 dependencies（仅为消费方解析
    // d.ts 中的保留式类型导出，见 README），而 tsup 默认把 dependencies 外置
    // ——没有本条，产物退化为 `import 'modern-xlsx'`，引擎泄漏为传递运行时
    // 依赖，两个 esbuild 插件在主线程链路失效（消费方 bundler 直接处理 raw
    // modern-xlsx，警告回潮）。运行时自包含因此只能靠本条强制打包。
    // 类型侧已核验（2026-09）：dts pass 无法内联 re-export 面，与依赖放
    // dependencies 还是 devDependencies 无关——modern-xlsx 是纯 exports
    // map 包（无顶层 types/main），tsup 的 dts resolver 解析不了它，
    // export-from 与 import+re-export 两种写法实测都外置为
    // `from 'modern-xlsx'`；抽取前 excel-exporter 2.6.9 的自包含 d.ts 靠的
    // 是手工复制的类型声明（其源码注释可证），对 core 的 49 符号 re-export
    // 面不可复制此法。依赖声明因此不能省：消费方 TS 需经它解析类型。
    noExternal: ["modern-xlsx"],
    splitting: true,
    treeshake: true,
    clean: true,
    sourcemap: true,
    target: "es2022",
    // copy-wasm 挂在 onSuccess 而非只在 "build" script：主配置 clean:true，
    // watch 模式首次构建即清空 dist（连带删掉上次 build 复制的 wasm）且
    // 此后不再回补，开着 pnpm dev 时 Node 自动初始化/集成测试会因
    // dist/modern-xlsx.wasm 缺失而失败。onSuccess 在普通构建与 watch 的
    // 每次重建后都执行，"build" 与 "dev" 两条链路由此统一回补。
    // （迁移自 excel-exporter 2.1.5 的修复，教训原因保留。）
    onSuccess: "node scripts/copy-wasm.mjs",
    // Browser resolution: without this, tsup defaults to platform "node" and
    // modern-xlsx's Node entries can drag Node-only shims into consumer
    // browser builds. See src/tsup-plugins.ts for the paired rewrites.
    platform: "browser",
  },
  {
    entry: { tsup: "src/tsup-plugins.ts" },
    format: ["esm"],
    // 不要 resolve：tsup 自身的 .d.ts 引用了 .cts 文件，rollup-dts 解析不了
    // （实测报错）。emit 保留式 `import type { Options } from "tsup"` 即可——
    // 该子路径只被其他包的 tsup.config.ts 消费，那里 tsup 必然是 devDep。
    dts: true,
    splitting: false,
    treeshake: true,
    sourcemap: true,
    target: "es2022",
    platform: "node",
  },
] as Options[]);
