# API：`createPptPreview`

```ts
import { createPptPreview } from "@marcusok/ppt-preview";

function createPptPreview(
  container: HTMLElement,
  options: PptPreviewOptions,
): PptPreviewInstance;
```

向 `container` 挂载只读 `.pptx` 预览并返回实例句柄。全部公开类型自包含于包内（不 re-export 引擎类型），见 `dist/index.d.ts`。

## `PptPreviewOptions`

| 选项                  | 类型                                        | 默认                      | 说明                                                                                                                                                                                                                                                                                                     |
| --------------------- | ------------------------------------------- | ------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `source`              | `File \| Blob \| Uint8Array \| ArrayBuffer` | —                         | **必填。** 字节在引擎运行前做防御性复制（解析器会 transfer 其输入 buffer）；调用方原始字节保持可用。                                                                                                                                                                                                     |
| `password`            | `string`                                    | —                         | MS-OFFCRYPTO Agile 加密文件的打开密码。解密在主线程执行（WebCrypto，需安全上下文，约 +1s）。ZIP 条目加密不覆盖。                                                                                                                                                                                         |
| `slideWidth`          | `number`                                    | `960`                     | 每页 CSS 像素宽度（小数向下取整）；高度按 deck 宽高比。必须是有限正数 —— 否则同步抛 `TypeError`。                                                                                                                                                                                                        |
| `dpr`                 | `number`                                    | `window.devicePixelRatio` | canvas 位图缩放。必须是有限正数 —— 否则同步抛 `TypeError`。                                                                                                                                                                                                                                              |
| `enableTextSelection` | `boolean`                                   | `true`                    | 文本层（选择/复制）。runs 与 canvas 绘制同一趟收集 —— 无双倍渲染。定位走 CSSOM；CSP 详见[边界说明](/zh/packages/ppt-preview/guide/02-limits)。                                                                                                                                                           |
| `useGoogleFonts`      | `boolean`                                   | `false`                   | 允许为缺失字体从 Google Fonts 拉取替代字形。默认关闭：不发起隐式外部请求。                                                                                                                                                                                                                               |
| `wasmUrl`             | `string \| URL`                             | 引擎锚点                  | WASM 资产 URL 的显式逃生门（自托管 / 资产路径被改写场景）。bundler 场景默认零配置（入口旁资产自动 emit）。                                                                                                                                                                                               |
| `styleNonce`          | `string`                                    | —                         | 作用于**本包**注入 `<style>` 的 CSP nonce（style-src）。CSSOM 写入的定位在 Chromium 下不受 style-src-attr 约束（见边界说明）。                                                                                                                                                                           |
| `renderOptions`       | `PptRenderOptions`                          | —                         | 白名单透传（见下）。                                                                                                                                                                                                                                                                                     |
| `onRendered`          | `(info) => void`                            | —                         | 由**首张成功渲染页**触发一次。`slideCount` 为原始总页数（含隐藏页）；`firstSlideMs` 是该页 canvas + 文本层耗时；`durationMs` 从 create 起算（读取、嗅探、解密、**全 deck preflight**、渲染）。全部页失败、空 deck，或创建时容器不可见/尺寸为 0（占位进不了渲染带）时不触发；容器可见后才开始渲染并触发。 |
| `onError`             | `(error) => void`                           | —                         | 加载失败与逐页渲染失败，归一为公开错误码。destroy 触发的取消静默。用户回调异常被隔离（`console.warn`），不会误报为预览错误。                                                                                                                                                                             |

### `PptRenderOptions`

| 选项              | 类型                                             | 默认     | 说明                                                                                                                                                               |
| ----------------- | ------------------------------------------------ | -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `cjkFallback`     | `"auto" \| "kr" \| "sc" \| "tc" \| "hk" \| "jp"` | `"auto"` | 缺失 CJK 字体的替代字形目标。                                                                                                                                      |
| `resourceLimits`  | `PptResourceLimits`                              | 引擎默认 | 压缩包预算：`maxArchiveEntryBytes`、`maxTotalInflatedBytes`、`maxArchiveEntries`（`number \| null`；`null` 关闭某项上限）。超预算报 `UNSUPPORTED` 而非 `CORRUPT`。 |
| `workerTimeoutMs` | `number`                                         | 引擎默认 | 解析 worker 单次响应超时；同时约束正常加载与加载期间 destroy 的最坏等待时间。                                                                                      |

引擎支持但本包有意**不**暴露的参数：`mode`（固定 `"main"`）、`progressiveLayout`（关闭 —— 页数与占位稳定）、`math`/`threeD`/`regionMap`/`chartEx`/`tiff`（可选渲染器，首版不注入）与已废弃的 `maxZipEntryBytes`（用 `resourceLimits.maxArchiveEntryBytes`）。

## `PptPreviewInstance`

### `destroy(): void`

幂等。断开观察器、作废在途渲染（引擎无单页 cancel 句柄 —— 完成路径清理）、canvas 置 0、移除注入 DOM、销毁引擎（终止解析 worker）。回调静默。加载期间调用时：旧加载作废，引擎实例到手即销毁（非立即取消 —— 上游 `load()` 无中止信号）。

### `goToSlide(index: number): void`

滚动到指定页（0-based、原始页序、含隐藏页）。识别到滚动容器时只滚该容器；容器本身不滚动（普通文档流）时回退滚动 `window`，外层页面会联动。

混合边界（有意设计）：

| 场景                        | 行为                                                 |
| --------------------------- | ---------------------------------------------------- |
| 非整数索引                  | 抛 `TypeError`（校验先于状态检查，destroy 后同样抛） |
| 加载完成前 / `destroy()` 后 | no-op                                                |
| 加载完成后越界              | 抛 `RangeError`                                      |

### `nextSlide(): void` / `prevSlide(): void`

从**当前最可见页**起算（按需测量 —— 手动滚动始终被尊重），而非最后一次 `goToSlide`。边界、加载前与 destroy 后为 no-op。

### `get slideCount(): number`

原始总页数（含隐藏页）。加载完成前、`destroy()` 后与空 deck 均为 `0`。

## 错误码

```ts
type PptPreviewErrorCode =
  | "UNSUPPORTED" // 不是 .pptx / 引擎不支持项 / 资源超预算
  | "LEGACY_FORMAT" // 老式二进制 .ppt
  | "PASSWORD_PROTECTED" // 加密：缺密码或密码错误（message 区分）
  | "CORRUPT" // 可稳定区分的结构损坏
  | "UNKNOWN"; // 源读取 / Worker / WASM / CSP / 资产 URL / 未知陷阱
```

引擎与源读取失败携带 `cause`（原始错误）供诊断；格式嗅探或 deck 尺寸非法在本包本地构造的错误可能不带 `cause`。
