import { defineConfig, type Options } from "tsup";
// 引擎 esbuild 插件自共享底层包复用（与 excel-exporter 同一模式）
import { rewriteWasmBgUrl, dropNodeFsPromises } from "@marcusok/xlsx-core/tsup";

const shared: Partial<Options> = {
  esbuildPlugins: [rewriteWasmBgUrl, dropNodeFsPromises],
};

// 与 excel-exporter 相同的双配置模式：
//  - 主入口 external @marcusok/xlsx-core（与导出包共享单份引擎与 wasm 的
//    关键），fflate（样式覆盖层用）bundle 进产物；
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
    external: ["@marcusok/xlsx-core"],
    // wasm 兜底转发：worker 默认按重写后的相对 URL 解析到本包 dist，
    // 主线程转发的显式 URL 不可用时仍能找到二进制（见 worker-client 注释）。
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
