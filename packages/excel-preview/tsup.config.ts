import { defineConfig, type Options } from "tsup";
// 引擎 esbuild 插件自共享底层包复用（与 excel-exporter 同一模式）。
// xlsx-core 已转为仓库私有包（不再发布 npm），仅作本包的构建期依赖。
import { rewriteWasmBgUrl, dropNodeFsPromises } from "@marcusok/xlsx-core/tsup";

const shared: Partial<Options> = {
  esbuildPlugins: [rewriteWasmBgUrl, dropNodeFsPromises],
};

// 与 excel-exporter 相同的双配置模式：
//  - 主入口把引擎层（@marcusok/xlsx-core，含 modern-xlsx）打进本包产物
//    （私有包，消费方只依赖本包；npm 发布面收敛为两个业务包）。代价是
//    同页同时使用两个业务包时主线程各持一份引擎与 wasm——业务后台通常
//    一页只用其一，详见文档站 package-relationships 指南。fflate（样式
//    覆盖层用）同样 bundle 进产物。注意 dts 侧无法内联 modern-xlsx 的
//    类型再导出（见 xlsx-core/tsup.config.ts 实测记录），dist/index.d.ts
//    保留 `from 'modern-xlsx'` 外部引用——package.json 的 dependencies
//    因此保留 modern-xlsx（仅类型用、精确钉打包同版）。
//  - worker 入口单文件自包含（browser module worker 不能解析裸导入），
//    core（含 modern-xlsx）与 fflate 一并打入。构建后
//    `grep -c "^import" dist/parse.worker.js` 必须为 0。
export default defineConfig([
  {
    ...shared,
    entry: { index: "src/index.ts" },
    format: ["esm"],
    dts: { resolve: true },
    splitting: true,
    treeshake: true,
    clean: true,
    sourcemap: true,
    target: "es2022",
    noExternal: ["@marcusok/xlsx-core"],
    // 引擎打进本包后，这份二进制是主线程 loader 默认 URL 与 worker 兜底
    // 相对 URL 的正主（worker 默认按重写后的相对 URL 解析到本包 dist，
    // 主线程转发的显式 URL 不可用时仍能找到二进制，见 worker-client 注释）。
    onSuccess: "node scripts/copy-wasm.mjs",
    platform: "browser",
  },
  {
    ...shared,
    entry: { "parse.worker": "src/worker/parse.worker.ts" },
    format: ["esm"],
    dts: false,
    splitting: false,
    treeshake: true,
    sourcemap: true,
    target: "es2022",
    platform: "browser",
    noExternal: ["@marcusok/xlsx-core", "fflate"],
    clean: false,
  },
]);
