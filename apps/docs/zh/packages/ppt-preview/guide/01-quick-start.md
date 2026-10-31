# 快速开始

## 安装

```bash
pnpm add @marcusok/ppt-preview
# npm i @marcusok/ppt-preview / yarn add @marcusok/ppt-preview 均可
```

本包仅 ESM，渲染引擎（Rust/WASM，`@silurus/ooxml` 0.89.6）在构建期打进包内 —— 消费方**零传递依赖**。Node ≥ 22。

> **Vite 5–7 dev server。** 生产构建（Vite 8+、webpack 5）会自动 emit 解析器 wasm；但 Vite 5–7 用 esbuild 把 npm 安装的依赖预打包进 `node_modules/.vite/deps/`。预打包副本会原样保留 `new URL("pptx_parser_bg.wasm", import.meta.url)`，wasm 请求因此落到 `.vite/deps/` 下（文件不存在，Vite 以 HTML fallback 或 404 响应），预览报 `UNKNOWN`。把包加入 `optimizeDeps.exclude` 并重启 dev server：
>
> ```ts
> // vite.config.ts
> export default { optimizeDeps: { exclude: ["@marcusok/ppt-preview"] } };
> ```
>
> 生产构建不受影响。（Vite 8 的 rolldown 预打包会正确改写该字面量；workspace 链接包不参与预打包。）

## 渲染一份演示文稿

```ts
import { createPptPreview } from "@marcusok/ppt-preview";

const container = document.querySelector<HTMLDivElement>("#preview");

const preview = createPptPreview(container!, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onRendered: ({ slideCount, firstSlideMs, durationMs }) => {
    console.log(
      `共 ${slideCount} 页；首张 ${firstSlideMs}ms，端到端 ${durationMs}ms`,
    );
  },
  onError: (e) => console.error(e.code, e.message), // 错误码见下
});
```

容器应当是真正会滚动的元素（`overflow: auto` 且有确定高度，如固定高度面板）。预览是垂直堆叠的等比占位；页面进入视口 ±1 屏才渲染，远离视口 ±3 屏外释放 canvas 显存。

## 尺寸

- `slideWidth`（默认 `960`）—— 每页的 CSS 像素宽度；高度按 deck 自身宽高比（16:9、4:3……）。非法值同步抛 `TypeError`。
- `dpr` —— 覆盖设备像素比（默认 `window.devicePixelRatio`）。高分屏下 canvas 位图放大、CSS 尺寸不变。

首版没有 resize 监听：容器尺寸变化时 `destroy()` 后重建实例。

## 页导航（0-based）

```ts
preview.goToSlide(0); // 滚到第一页
preview.goToSlide(count - 1);
preview.nextSlide(); // 从当前最可见页起算
preview.prevSlide();
preview.slideCount; // 加载完成前 / destroy 后为 0
```

混合边界 —— 有意设计，请仔细阅读：

- 非整数索引**抛 `TypeError`**（编程错误，fail-fast）。
- 加载完成前或 `destroy()` 后**no-op**（占位尚不存在）。
- 加载完成后越界索引**抛 `RangeError`**。

隐藏页全部渲染并计数；索引遵循原始 deck 页序（不跳页）。

## 加密文件

```ts
const preview = createPptPreview(container!, {
  source: encryptedFile,
  password: "打开密码", // PowerPoint「打开密码」（MS-OFFCRYPTO Agile）
  ...
});
```

解密经主线程 WebCrypto 执行 —— 需要安全上下文（HTTPS 或 localhost），通常在解析前增加约 1s。缺密码或密码错误报 `PASSWORD_PROTECTED`（message 区分两种情况）。ZIP 条目加密不在覆盖范围。

## 错误码

| 错误码               | 含义                                                                     |
| -------------------- | ------------------------------------------------------------------------ |
| `UNSUPPORTED`        | 不是 `.pptx`（无法识别的容器），或密码无法解决的不支持项（含资源超预算） |
| `LEGACY_FORMAT`      | 老式二进制 `.ppt` —— 请另存为 `.pptx`                                    |
| `PASSWORD_PROTECTED` | 加密文件：缺密码或密码错误（message 区分）                               |
| `CORRUPT`            | 结构损坏（仅在可稳定区分时使用）                                         |
| `UNKNOWN`            | 源读取、Worker/WASM 启动、CSP 或资产 URL 失败 —— 文件本身没有问题        |

## 销毁

```ts
preview.destroy();
```

幂等。断开观察器、作废在途渲染、canvas 置 0、移除注入 DOM 并终止引擎解析 Worker。destroy 后所有回调静默。若在引擎加载期间调用：旧结果作废，引擎实例到手后立即销毁（上游 `load()` 无中止信号 —— 这不是立即取消）。

## SSR

模块可以在服务端 import，但 `createPptPreview` 只能在具备 `Worker` 与 WebCrypto 的浏览器环境执行（在 `onMounted` / `useEffect` 内守卫）。
