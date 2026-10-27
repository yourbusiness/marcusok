/**
 * 资产 URL 推导：worker / cmaps / standard_fonts 三项随包分发（dist/），
 * wasm / iccs 不分发（上游缺失时优雅降级）。
 *
 * worker 以单文件静态字面量为唯一锚点——bundler 资产管道（Vite asset
 * modules / webpack 5）只识别 `new URL("<file>", import.meta.url)` 的
 * 字面量原文并 emit 为资产（同仓 excel-preview 的 wasm/worker 已实证：
 * VitePress build 下自动 emit 为 base-aware hashed assets）；play 的
 * rewriteAssetLiterals 也按该原文逐字匹配。**不要**把字面量改成变量拼接
 * 或换 base——两套机制都会失配。
 *
 * src 直跑形态（play 源码联调等模块 URL 落在 src/ 下的场景，且未经
 * 字面量改写覆盖）锚点会指向 src/ 下不存在的文件——用 `../dist/` 运行时
 * 兜底。兜底刻意写成变量拼接而非静态字面量：静态 `../` 字面量会被
 * bundler 的资产分析捕获，而逐文件改写机制不覆盖 `../` 前缀，改写不了
 * 反而干扰 emit（vite 8/rolldown 对该形态的实测教训见 play 的
 * src-asset-overrides.ts 头注释）。
 *
 * cmaps / standard_fonts 从 worker URL 推导目录（worker 与两个资产目录
 * 在包 dist 内同层）。显式 `workerSrc` 允许写成根相对/相对路径（自托管
 * 最常见写法），推导前先按页面基址解析成绝对 URL——`URL` 构造器的 base
 * 不接受相对 URL，直接拿 `/assets/x.mjs` 当 base 会抛 `Invalid base URL`。
 * **已知边界**：bundler 把 worker emit 为 `<assets>/pdf.worker-<hash>.mjs`
 * 后，推导得到 `<assets>/cmaps/`——资产化目录里没有 cmaps，CID 文档
 * （中文 PDF）会静默劣化。bundler 场景的正确姿势是显式传 `cMapUrl` /
 * `standardFontDataUrl`（docs 站 demo 即走此路）或配 vite-static-copy
 * 一类插件，见 limits 文档的场景矩阵。
 */

export interface PdfAssetUrls {
  workerSrc: string;
  /** 以 `/` 结尾的目录 URL。 */
  cMapUrl: string;
  /** 以 `/` 结尾的目录 URL。 */
  standardFontDataUrl: string;
}

/**
 * 把显式 workerSrc 解析成绝对 URL：它可能是绝对 URL、根相对路径或相对
 * 路径。`URL` 构造器的 base 本身必须是绝对 URL，而测试环境（happy-dom）
 * 的 document.baseURI 是 `about:blank`——逐级尝试 document.baseURI /
 * location.href / localhost 兜底，避免测试与特殊宿主环境抛
 * `Invalid base URL`。
 */
function absoluteUrl(url: string): URL {
  try {
    return new URL(url);
  } catch {
    // 相对路径：继续尝试基址
  }
  const bases: string[] = [];
  if (typeof document !== "undefined" && document.baseURI) {
    bases.push(document.baseURI);
  }
  if (typeof location !== "undefined" && location.href) {
    bases.push(location.href);
  }
  bases.push("http://localhost/");
  for (const base of bases) {
    try {
      return new URL(url, base);
    } catch {
      // 该基址不可用（如 about:blank），尝试下一个
    }
  }
  throw new TypeError(`Cannot resolve asset URL: ${url}`);
}

function resolveWorkerUrl(): string {
  // 只在真正的源码直跑形态（本文件位于 <pkg>/src/assets.ts）走 ../dist/
  // 兜底。收紧到精确文件名，避免生产 bundle 路径恰好包含 /src/ 时误覆盖
  // bundler 已 emit 的 worker URL。
  if (/\/src\/assets\.(?:ts|js)$/.test(new URL(import.meta.url).pathname)) {
    // 运行时拼接（非静态字面量，见头注释）：src 直跑时 ../dist/ 才是
    // 资产真实位置
    const prefix = "../dist/";
    return new URL(prefix + "pdf.worker.min.mjs", import.meta.url).href;
  }
  // 唯一锚点字面量：保持原文，勿改写成变量拼接
  return new URL("./pdf.worker.min.mjs", import.meta.url).href;
}

/** 从 worker URL 推导同层资产目录（相对 URL 先按页面基址绝对化）。 */
function siblingDir(
  workerUrl: string,
  dir: "cmaps" | "standard_fonts",
): string {
  return new URL(`${dir}/`, absoluteUrl(workerUrl)).href;
}

export function resolveAssetUrls(explicit: {
  workerSrc?: string;
  cMapUrl?: string;
  standardFontDataUrl?: string;
}): PdfAssetUrls {
  const workerSrc = explicit.workerSrc ?? resolveWorkerUrl();
  return {
    workerSrc,
    cMapUrl: explicit.cMapUrl ?? siblingDir(workerSrc, "cmaps"),
    standardFontDataUrl:
      explicit.standardFontDataUrl ?? siblingDir(workerSrc, "standard_fonts"),
  };
}
