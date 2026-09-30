# @marcusok/excel-preview

浏览器端 xlsx 只读预览，与 [@marcusok/excel-exporter](/zh/packages/excel-exporter/) 共用同一 WASM 引擎：解析在 Web Worker 内完成，渲染是框架无关的 DOM 虚拟滚动网格，配套补偿层还原 Excel 实际的数字格式与主题色。

## 能力清单

| 能力         | 说明                                                                                                                       |
| ------------ | -------------------------------------------------------------------------------------------------------------------------- |
| 文件格式     | `.xlsx` / `.xlsm`（ZIP/OOXML）、加密工作簿（`password` 选项）、`.csv`（UTF-8 / GB18030、分隔符嗅探）                       |
| Worker 解析  | 解析与模型构建全部在共享 Web Worker 内（实测 10 万 × 10 格约 1.5s，全程不冻结 UI）                                         |
| 虚拟滚动     | DOM 渲染器只挂载视口内格子（外加缓冲）；跨视口的合并单元格按覆盖区整体召回                                                 |
| 布局还原     | 列宽、行高、隐藏行列、合并单元格、冻结窗格（四象限分层）、sheet 页签（隐藏表不出现在页签）                                 |
| 样式还原     | 字体、纯色/图案/渐变填充、边框、对齐（换行/缩进/旋转）、文件声明的网格线开关                                               |
| 主题色找回   | 自解 `styles.xml` + `theme1.xml` 的覆盖层——引擎读取侧会把 theme/tint/indexed 颜色丢弃为 null                               |
| 数字格式忠实 | 内置 id 用 Excel 实际行为表（非 ECMA 标准串）、负号/会计括号/货币字面量补偿、累计时长 `[h]:mm:ss`、分钟邻接、date1904 平移 |
| 错误归一     | 携带错误码的友好报错：`PASSWORD_PROTECTED`、`LEGACY_FORMAT`（旧版 .xls）、`CORRUPT`、`UNSUPPORTED`、`WASM`、`UNKNOWN`      |

## 安装

```bash
pnpm add @marcusok/excel-preview
```

一条安装命令即得全部：解析引擎（modern-xlsx）在构建期打包进本包 `dist`，WASM 二进制通过本包自己的 `exports` 映射分发。详见 [快速上手](/zh/packages/excel-preview/guide/01-quick-start)。

## 最小示例

```ts
import { createPreview } from "@marcusok/excel-preview";

const preview = createPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onParsed: (info) => console.log(info.sheetNames, info.duration),
  onError: (e) => console.error(e.code, e.message),
});

// setSheet()/getSheetNames() 在解析完成前是空操作；
// destroy() 立即生效（阻断尚未完成的渲染与回调）
preview.setSheet("Sheet2"); // 名称或索引
preview.destroy(); // 卸载并释放资源
```

低层 API `parseWorkbookBytes(bytes, { password })` 返回纯数据模型（`PreviewWorkbook`，结构化克隆安全），供 React/Vue 薄封装、SSR 或自研渲染器消费。

<ClientOnly>
<PreviewDemo />
</ClientOnly>

## v1 不做的事

- 编辑与公式重算——公式单元格渲染**缓存值**（业界共识；无缓存值的公式渲染为空）。
- 旧版 `.xls`（BIFF8）与 `.ods`——明确报错拒绝。
- 图表、图片、形状、条件格式、富文本 runs（降级为拼接纯文本）、外部超链接 URL（引擎读取侧丢弃；链接样式仍会显示）。
