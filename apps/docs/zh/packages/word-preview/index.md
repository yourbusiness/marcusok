# @marcusok/word-preview

浏览器端 `.docx` 只读预览，视觉保真——分页、页眉页脚、脚注尾注、批注与修订按 Word 存盘的样子渲染。渲染引擎为 [docx-preview](https://github.com/VolodymyrBaydalka/docxjs)（构建期打进本包 `dist`，消费方零传递依赖），本包在其上叠加格式嗅探、可读错误码、安全加固默认值与实例级样式隔离。

## 能力清单

| 能力       | 说明                                                                                                                                                                                  |
| ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 文件格式   | `.docx` / `.docm`（ZIP/OOXML）。`.doc` 老格式与加密 docx 同为 OLE2 魔数，统一报 `LEGACY_FORMAT`；罕见的 zip 直接加密形态报 `PASSWORD_PROTECTED`                                       |
| 视觉保真   | 分页符（显式 `<w:br type="page"/>`、`lastRenderedPageBreak` 标记、分节符）、页眉页脚、脚注尾注、批注、修订、内嵌图片与字体——由 docx-preview 引擎渲染                                  |
| 格式嗅探   | 魔数 + zip 加密位 + 必需部件校验（`[Content_Types].xml`、`_rels/.rels`、`word/document.xml`）在引擎**之前**执行，非 Word 的 zip 得到可读的 `UNSUPPORTED` 报错而非莫名其妙的空引用堆栈 |
| 加固默认值 | altChunk（引擎渲染为无 sandbox 的 `srcdoc` iframe）默认**关闭**；渲染后对 `http`/`https`/`mailto`/`tel` 白名单之外的超链接置空 href（保留链接文本）                                   |
| 多实例     | 每个实例在唯一样式类名下渲染，同页两份文档的编号/主题 CSS 互不串扰（引擎固定类名会互相覆盖）                                                                                          |
| CSP 支持   | `styleNonce` 以"换节点"方式给注入的 `<style>` 加 nonce（CSP 在插入时拦截，事后补属性无效）；内联 style 属性与 `data:` 资源的要求见 [限制](/zh/packages/word-preview/guide/02-limits)  |
| 取消       | `destroy()` 同步摘除内部 wrapper——迟到的渲染结果写入已脱离的节点后被丢弃，所有回调静默。取消不是错误（不触发 `onError`）                                                              |
| 错误归一   | 携带错误码的报错：`UNSUPPORTED`、`LEGACY_FORMAT`、`PASSWORD_PROTECTED`、`CORRUPT`、`UNKNOWN`                                                                                          |

## 安装

```bash
pnpm add @marcusok/word-preview
```

一条安装命令即得全部：docx-preview 与其 zip 依赖（jszip）在构建期打进本包 `dist`（二者均为 devDependencies——npm 消费方**零传递依赖**）。详见 [快速上手](/zh/packages/word-preview/guide/01-quick-start)。

## 最小示例

```ts
import { createWordPreview } from "@marcusok/word-preview";

const preview = createWordPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onRendered: (info) => console.log(info.pageCount, info.durationMs),
  onError: (e) => console.error(e.code, e.message),
});

preview.destroy(); // 卸载 + 兜住任何在途渲染
```

有意**不提供低层数据模型 API**：引擎仅保证 `renderAsync` 稳定（上游 `WordDocument` 是 opaque 的 `any` stub），无法诚实地提供 `parse()` 式接口。这一点与提供 `parseWorkbookBytes` 的 [@marcusok/excel-preview](/zh/packages/excel-preview/) 不同，该模式见其文档。

<ClientOnly>
<WordPreviewDemo />
</ClientOnly>

## 不在范围内

- **实时排版分页**——只在显式分页/分节符与编辑器存盘的 `lastRenderedPageBreak` 标记处断页；引擎不做文本重排。
- **TOC 字段**——目录字段不渲染。
- **编辑**——这是预览器，不编辑、不另存。
- **Worker 解析**——引擎在主线程产出 DOM（上游约束）；大文档的 DOM 构建期间会阻塞 UI（见 [限制](/zh/packages/word-preview/guide/02-limits)）。
