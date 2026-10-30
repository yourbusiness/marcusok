/**
 * WASM 资产 URL 解析：透传优先，包内 dist 锚点兜底。
 *
 * 上游 chunk 以 `new URL("pptx_parser_bg.wasm", import.meta.url)` 定位
 * wasm，bundler 消费方（Vite 等）在包 dist 内检测到该文件即 emit 并重写
 * URL（spike S1 实测）——**默认 bundler 场景零配置**。wasmUrl 选项是自
 * 托管/资产路径被改写场景的显式逃生门，不做多级目录推导（与 pdf-preview
 * 的 worker/cmaps 三级推导不同：pptx 只有一个必需资产，复杂推导没有收益）。
 *
 * src 直跑（play 把本包主入口别名到 src 联调）场景下模块 URL 精确匹配
 * `src/assets.ts`，资产实际在包 dist/——回退 `../dist/`。匹配必须精确到
 * 模块文件名，不能用路径包含 `/src/` 的宽松判断，否则会误伤消费方产物。
 */

/** 解析 wasm 资产 URL：显式选项 → 包内锚点（bundler emit）→ ../dist 兜底。 */
export function resolveWasmUrl(
  explicit: string | URL | undefined,
  /**
   * 锚点基址（默认本模块 URL）。仅测试注入非 src 锚点以覆盖 bundler
   * emit 分支使用；生产调用方不传。
   */
  anchorBase: string | URL = import.meta.url,
): string | URL | undefined {
  if (explicit !== undefined) {
    // URL 实例直接透传（上游 LoadOptions 接受 string | URL）
    if (explicit instanceof URL) return explicit;
    if (explicit !== "") return explicit;
    throw new TypeError("Invalid wasmUrl: pass a non-empty string or URL.");
  }
  // 不传时不干预上游默认（new URL 锚点）；仅当本模块 URL 明确落在源码目录
  // （src 直跑场景，非 bundler emit）时补 ../dist/ 兜底。检测精确到模块
  // 文件名（对齐 pdf-preview 的 resolveWorkerUrl）：宽松的 /\/src\// 会在
  // 消费方 bundle 产物路径恰含 /src/ 段（自定义 outDir、站点路由等）时
  // 误覆盖 bundler 已 emit 的 wasm URL → 404 且极难排查。
  const moduleUrl = new URL(anchorBase);
  if (/\/src\/assets\.(?:ts|js)$/.test(moduleUrl.pathname)) {
    return new URL("../dist/pptx_parser_bg.wasm", moduleUrl);
  }
  return undefined;
}
