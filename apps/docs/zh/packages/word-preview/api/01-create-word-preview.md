# createWordPreview

唯一的高层入口。把 `.docx` 渲染进容器，附带格式嗅探、错误归一、加固默认值与渲染后加固。

> **命名说明：** `@marcusok/excel-preview` 导出的是 `createPreview`；本包有意用 `createWordPreview`——同页同时 import 两个包时不歧义。

```ts
import { createWordPreview } from "@marcusok/word-preview";

function createWordPreview(
  container: HTMLElement,
  options: WordPreviewOptions,
): WordPreviewInstance;
```

## WordPreviewOptions

| 选项             | 类型                                        | 默认值  | 说明                                                                                                                                                                                                                                                      |
| ---------------- | ------------------------------------------- | ------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `source`         | `File \| Blob \| Uint8Array \| ArrayBuffer` | —       | 文档字节。                                                                                                                                                                                                                                                |
| `renderOptions`  | `WordRenderOptions`                         | —       | 白名单引擎选项（见下）。                                                                                                                                                                                                                                  |
| `styleNonce`     | `string`                                    | —       | 注入 `<style>` 的 CSP nonce。以"换节点"实现（CSP 在插入时拦截——事后补属性无效）。见 [限制 → CSP](/zh/packages/word-preview/guide/02-limits#csp-要求)。                                                                                                    |
| `allowAltChunks` | `boolean`                                   | `false` | 渲染 altChunk（内嵌 HTML）。默认关闭——引擎用无 sandbox 的 `srcdoc` iframe。仅限受信文档；渲染后会补 `sandbox=""` 属性（残余竞态见 [限制](/zh/packages/word-preview/guide/02-limits)）。                                                                   |
| `onRendered`     | `(info) => void`                            | —       | 每次渲染成功触发一次。`info.pageCount` 按 `section` 元素计数（仅在默认分页选项下有意义）；`info.durationMs` 从 `createWordPreview` 调用起计——涵盖源读取、格式嗅探与**含图片/字体加载**的完整渲染（引擎 resolve 前 await 全部资源任务），不是纯 DOM 构建。 |
| `onError`        | `(error) => void`                           | —       | 致命错误。destroy 触发的取消是静默的，永不进入该回调。                                                                                                                                                                                                    |

错误带 `code`：

| 错误码               | 含义                                                                                                                  |
| -------------------- | --------------------------------------------------------------------------------------------------------------------- |
| `UNSUPPORTED`        | 非 zip 容器，或 zip 缺 Word 必需部件（`[Content_Types].xml`、`_rels/.rels`、`word/document.xml`）——即不是 Word 文档。 |
| `LEGACY_FORMAT`      | OLE2 容器：`.doc` 老格式**或**加密 docx（同魔数；两者均不支持）。                                                     |
| `PASSWORD_PROTECTED` | zip general purpose 加密位置位（罕见的直接加密形态）。                                                                |
| `CORRUPT`            | zip 结构损坏，或必需部件齐全但解析/渲染失败（`cause` 保留原始错误）。                                                 |
| `UNKNOWN`            | 其他（如源 `Blob` 读取失败）。                                                                                        |

## WordRenderOptions

引擎的 20 个选项拆成 **13 项可透传** + 7 项由本包接管。传入接管键无效——会被覆盖。

**可透传：**

`breakPages`、`inWrapper`、`hideWrapperOnPrint`、`ignoreWidth`、`ignoreHeight`、`ignoreFonts`、`ignoreLastRenderedPageBreak`（本包默认 `false`、引擎默认 `true`——设 `true` 恢复引擎行为）、`renderHeaders`、`renderFooters`、`renderFootnotes`、`renderEndnotes`、`renderComments`、`renderChanges`

全部为可选布尔值；未设置的键用引擎默认。`onRendered` 的 `pageCount` 仅在默认分页选项（`breakPages: true`、`inWrapper: true`）下有意义。

**接管项（不在类型里，共 7 项）：**

| 键                                                    | 值                    | 原因                                                            |
| ----------------------------------------------------- | --------------------- | --------------------------------------------------------------- |
| `className`                                           | `docx-<instanceSeq>`  | 每实例唯一——隔离引擎按文档注入的全局 CSS。                      |
| `useBase64URL`                                        | `true`                | 引擎从不 revoke blob URL；base64 规避泄漏（内存代价见限制页）。 |
| `renderAltChunks`                                     | 来自 `allowAltChunks` | 安全默认值；见上。                                              |
| `h` / `experimental` / `debug` / `trimXmlDeclaration` | —                     | 实验 hook 与内部细节——不进公开面。                              |

## WordPreviewInstance

```ts
interface WordPreviewInstance {
  destroy(): void;
}
```

`destroy()` 同步摘除内部 wrapper（用户容器保持纯挂载点）、丢弃任何在途渲染结果、静默所有回调。引擎没有取消 API（上游 0 处 abort 路径）——在途渲染的同步 DOM 构建段无法中断，但其输出会被丢弃。幂等；容器可立即被新实例复用。
