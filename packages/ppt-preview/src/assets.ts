/**
 * WASM 资产 URL 解析：透传优先，包内 dist 锚点兜底。
 *
 * 上游 chunk 以 `new URL("pptx_parser_bg.wasm", import.meta.url)` 定位
 * wasm，bundler 消费方（Vite 等）在包 dist 内检测到该文件即 emit 并重写
 * URL（spike S1 实测）——**默认 bundler 场景零配置**。wasmUrl 选项是自
 * 托管/资产路径被改写场景的显式逃生门，不做多级目录推导（与 pdf-preview
 * 的 worker/cmaps 三级推导不同：pptx 只有一个必需资产，复杂推导没有收益）。
 *
 * src 直跑（play 把本包主入口别名到 src 联调）场景下锚点落在 src/，资产
 * 实际在包 dist/——回退 `../dist/`。注意兜底分支以 anchor 变量为 base 构造
 * URL（而非并列的静态字面量）：源码里只保留一个 `./pptx_parser_bg.wasm`
 * 静态字面量，play 的 srcAssetOverrides 改写机制按字面量精确匹配，出现
 * 第二个 `../dist/` 字面量会漏改写并触发运行时相对解析的不确定行为。
 */

/** 解析 wasm 资产 URL：显式选项 → 包内锚点（bundler emit）→ ../dist 兜底。 */
export function resolveWasmUrl(
  explicit: string | URL | undefined,
): string | URL | undefined {
  if (explicit !== undefined) {
    // URL 实例直接透传（上游 LoadOptions 接受 string | URL）
    if (explicit instanceof URL) return explicit;
    if (explicit !== "") return explicit;
    throw new TypeError("Invalid wasmUrl: pass a non-empty string or URL.");
  }
  // 不传时不干预上游默认（new URL 锚点）；仅当锚点路径明显落在源码目录
  // （非 bundler emit 场景）时补 ../dist/ 兜底
  const anchor = new URL("./pptx_parser_bg.wasm", import.meta.url);
  if (/\/src\//.test(anchor.pathname)) {
    return new URL("../dist/pptx_parser_bg.wasm", anchor);
  }
  return undefined;
}
