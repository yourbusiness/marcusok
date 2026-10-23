# 限制与安全说明

本包是 [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) 引擎的薄加固封装。引擎的约束原样继承；本页逐项如实列出，连同封装层自身的加固与残余风险。

## 渲染

- **主线程渲染。** 引擎在主线程构建 DOM、无逐元素让出机制；大文档的 DOM 构建期间阻塞 UI。没有 Worker 模式（引擎内部模型不保证结构化克隆安全——上游约束，非取舍）。真实 Chromium 实测基线（6 轮，2026-10-22）：1000 段落 + 100×5 表格完整渲染 16–28ms（中位约 20ms，含 `createWordPreview` 全流程）；实际耗时随文档规模变化。
- **无实时排版分页。** 只在显式分页符（`<w:br type="page"/>`）、分节符与编辑器存盘的 `lastRenderedPageBreak` 标记处断页。引擎不做文本重排，渲染出的页数可能与 Word 实际计算的页数不同。本包默认 `ignoreLastRenderedPageBreak: false`（遵循存盘标记；引擎自身默认忽略）。
- **TOC 字段不渲染。**
- **无输入大小上限。** docx 是不可信输入；首版没有源字节/解压大小的可配置限制（无 zip 炸弹防护）。只预览你本来就会信任打开的文件。
- **容器必须在主文档内。** 渲染进 iframe 文档或 ShadowRoot 未验证、首版不支持（引擎读取全局 `document`、`Range`、`getComputedStyle`）。

## 文件格式

- **仅 `.docx` / `.docm`。** `.doc` 老格式（OLE2）不支持。
- **加密文件。** 主流"加密 docx"（Agile AES）整体是 OLE2 容器——与 `.doc` 同魔数；两者统一报 `LEGACY_FORMAT`，message 明确说明。罕见的 zip 直接加密形态由 zip general purpose 加密位检测，报 `PASSWORD_PROTECTED`。

## 安全加固（封装层做了什么）

- **altChunk 默认关闭。** 引擎把 altChunk（内嵌 HTML 部件）渲染进 `<iframe srcdoc>` 且**不带 `sandbox` 属性**，引擎自身默认还是 `renderAltChunks: true`——不可信文档可在渲染时执行脚本。本包把默认值翻转为 `false`。经 `allowAltChunks: true` 显式开启仅限受信文档；封装层会在渲染后补 `sandbox=""`——注意这是渲染后的属性赋值，与 iframe 首帧加载的竞态不保证在所有浏览器都赢。iframe 内部的内容也不在下述超链接白名单的覆盖范围内。
- **超链接 scheme 白名单。** 引擎把外链关系 target **原样直赋** href、无任何 scheme 校验——恶意 docx 可产出 `<a href="javascript:…">`。渲染完成后，封装层用 URL API 解析每个 `a[href]`（正确处理大小写、前导空白/控制字符、`//host` 形态），仅放行 `http:`、`https:`、`mailto:`、`tel:` 与锚点/相对形态；其余摘除 href（保留链接文本）。没有关闭开关——文档预览里想不出放行 `javascript:` 链接的正当场景。
- **引擎之前先嗅探格式。** 魔数、zip 加密位与必需部件检查先行，恶意/畸形输入得到可读错误（`UNSUPPORTED` / `CORRUPT`），而不是引擎的空引用堆栈。

## 残余风险（封装层没做什么）

- **CSS 注入。** 文档可控字符串（编号文本、字体名）会进入注入的 `<style>` CSS 文本。封装层的实例级类名隔离的是**合法**样式在实例间的串扰，恶意 payload 仍可闭合规则注入全局选择器（如 `body { … }`）——已在真实 Chromium 验证。现代浏览器不会因此执行脚本；影响是 UI 伪装与外部 `@import`/`url()` 请求。用 CSP（`style-src`、`img-src`、`font-src`）缓解。不做 CSS 值净化——那会破坏引擎合法的 base64 图片/字体。
- **实例间同名嵌入字体。** `@font-face` 按名字是全局的；两份文档嵌入同名但不同内容的字体——后渲染者胜出。

## CSP 要求

`styleNonce` 只覆盖注入的 `<style>` **元素**。引擎还大量使用内联 style **属性**（页面宽高/边距、图片定位），且 `useBase64URL: true` 后图片与字体是 `data:` URL。严格 CSP 因此需要全部三条：

- `style-src 'nonce-…'`（或 `'unsafe-inline'`）——注入的样式
- `style-src-attr 'unsafe-inline'`（或整体 `style-src 'unsafe-inline'`）——内联属性
- `img-src data:` 与 `font-src data:`——base64 资源

## 内存

`useBase64URL: true`（锁定）意味着图片与内嵌字体以 base64 形式驻留 DOM——比原始字节约 +33%，预览挂载期间常驻。图片密集的文档比 blob URL 方案更重；该取舍规避了引擎不 revoke 的 blob URL 泄漏。

## SSR

在 Node 里 import 本包是安全的（模块顶层不触碰任何 DOM 全局——已实测）。只有 `createWordPreview` 调用需要 DOM；在客户端调用即可。

## 多实例

每个实例在唯一的 `docx-<n>` 类名下渲染，隔离引擎按文档注入的 CSS（默认样式、文档样式、编号、主题变量）。残留：上述 `@font-face` 同名冲突。已在真实浏览器验证——两份编号不同的文档各自保持自己的编号。一个边缘注意项：实例计数器属于**已加载的库副本**——同页加载两份 `@marcusok/word-preview`（如两个版本并存）时各自从 1 重新计数、类名可能撞名；该形态不支持。
