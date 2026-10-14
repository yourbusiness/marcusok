# createPreview

高层入口：一次调用完成解析 + 渲染。

```ts
import { createPreview } from "@marcusok/excel-preview";

function createPreview(
  container: HTMLElement,
  options: PreviewOptions,
): PreviewInstance;
```

## PreviewOptions

| 选项            | 类型                                        | 默认值           | 说明                                                                                                                                                                                                                                       |
| --------------- | ------------------------------------------- | ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `source`        | `File \| Blob \| Uint8Array \| ArrayBuffer` | —（必填）        | 文件字节                                                                                                                                                                                                                                   |
| `password`      | `string`                                    | —                | 加密工作簿的密码（Agile AES-256）                                                                                                                                                                                                          |
| `sheet`         | `string \| number`                          | 文件 `activeTab` | 初始 sheet（名称或 0 起索引）。无效值（名称不存在、索引越界或非整数）会经 `onError` 报出 `Sheet not found: …`——与实例方法 `setSheet()` 同文案同码——随后仍回退渲染文件 `activeTab`（该值本身亦经归一化，见模型参考中的 `activeSheetIndex`） |
| `showHeaders`   | `boolean`                                   | `true`           | 行列表头（A/B/C + 1/2/3）                                                                                                                                                                                                                  |
| `showGridLines` | `boolean`                                   | 遵循文件         | 网格线                                                                                                                                                                                                                                     |
| `showTabs`      | `boolean`                                   | `true`           | sheet 页签栏（隐藏表永不出现）                                                                                                                                                                                                             |
| `onParsed`      | `(info: PreviewParsedInfo) => void`         | —                | 渲染就绪回调：首次解析渲染完成后与**每次 sheet 切换完成后**都会触发（`duration.parse` 复用首次解析耗时，`duration.total` 则一直从首次加载起累计——要衡量本次切换本身的耗时请用 `duration.render`）；上报 sheet 列表、规模与耗时             |
| `onError`       | `(error: PreviewError) => void`             | —                | 失败回调（错误码见下）                                                                                                                                                                                                                     |

`PreviewParsedInfo`：

```ts
interface PreviewParsedInfo {
  sheetNames: string[];
  sheetCount: number;
  rowCount: number; // 当前 sheet
  colCount: number;
  duration: {
    parse: number; // 毫秒；仅首次解析——后续切换 sheet 复用该值
    render: number; // 毫秒；本次渲染
    total: number; // 毫秒；自 createPreview 启动起累计，不是本次切换耗时
  };
}
```

## PreviewError 错误码

| 错误码               | 含义                                                                                                                                                                                                                                                                                                                                                                                                                                |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PASSWORD_PROTECTED` | 加密工作簿，`password` 缺失或错误                                                                                                                                                                                                                                                                                                                                                                                                   |
| `LEGACY_FORMAT`      | 旧版 `.xls`（BIFF8）——请另存为 `.xlsx`                                                                                                                                                                                                                                                                                                                                                                                              |
| `CORRUPT`            | ZIP 但不是有效 xlsx：部件缺失/损坏（`.ods` / `.docx` 同形包、截断文件），或工作簿一个 sheet 都没有                                                                                                                                                                                                                                                                                                                                  |
| `UNSUPPORTED`        | 既不是 ZIP 也不是可识别的纯文本——含 XML/HTML 伪表格（SpreadsheetML 2003、HTML 表格另存为 `.xls`）与 UTF-16 编码的 CSV                                                                                                                                                                                                                                                                                                               |
| `WASM`               | 当前环境不支持 WebAssembly，或引擎加载失败（资产 404、CSP 禁止 WebAssembly、网络失败）                                                                                                                                                                                                                                                                                                                                              |
| `UNKNOWN`            | 其他——含解析 worker 超时，其报文为专用的 `parse worker timed out after …ms (worker terminated); retry the parse`（worker 会被终止，且**不会**转到主线程重跑）；底层错误信息原样透传（失败来自预览自身 boot 路径时，原始错误对象在 `.cause`）                                                                                                                                                                                        |
| `ABORTED`            | 调用方主动取消——在途预览的 `destroy()`，或传给 `parseWorkbookBytes` 的 `signal` 被 abort。取消**不会**转到主线程重跑（对已取消的输入复跑违背取消意图）。`createPreview` 对该码静默：**不**触发 `onError`。语义边界：主线程路径在下一个 `await` 边界放弃后续大计算（同步 WASM 段无法中途打断）；worker 路径调用方立即解除等待、迟到的 worker 结果经已清空的 pending 槽位丢弃——取消**不会** terminate 共享 worker（那会拒绝兄弟请求） |

错误对象的完整形状：

```ts
interface PreviewError {
  code: PreviewErrorCode;
  message: string;
  cause?: unknown; // 原始错误，仅 boot 路径异常时携带
}
```

`cause` 只由 boot 路径的异常兜底附带；解析管线上报的解析失败只有 `code` + `message`、没有 `cause`——不要把诊断逻辑建立在它总是存在之上。

## PreviewInstance

```ts
interface PreviewInstance {
  destroy(): void; // 卸载 DOM、释放资源
  setSheet(nameOrIndex: string | number): void; // 切换 sheet
  getSheetNames(): string[]; // 全部 sheet，文件顺序（含隐藏表）
}
```

`destroy()` 移除渲染 DOM，**并中止在途解析**（上表的 `ABORTED` 码——boot 路径静默返回，取消的预览不会触发 `onError`/`onParsed`）。解析 worker 是模块级共享资源（与导出包一致，有意跨实例保温）；解析完成后不持有任何文件数据，取消也不会 terminate 它（迟到的结果直接丢弃）。

隐藏 sheet 语义：`setSheet()` 按名称或 0 起索引查找，**不过滤隐藏 sheet**——与页签栏（永不显示它们）不同，显式 `setSheet("HiddenSheet")` 会渲染它。无效的名称/索引经 `onError` 报出（`UNKNOWN` 码）。`getSheetNames()` 同样包含 hidden / veryHidden 的 sheet，按文件顺序——若用它驱动自己的 sheet 选择器，请自行过滤。

## parseWorkbookBytes

低层解析，不做 DOM：

```ts
import { parseWorkbookBytes } from "@marcusok/excel-preview";

const workbook = await parseWorkbookBytes(bytes, {
  password: "…",
  signal, // 可选 AbortSignal——abort 时以 ABORTED 码拒绝
});
```

失败时抛出带 `.code` 属性的 `Error`（错误码同上）；`signal` 中止时抛出 `ABORTED` 码。模型字段见[数据模型](/zh/packages/excel-preview/api/02-model)。

## 其他导出

| 导出              | 说明                                                                                                                                                                                                                                    |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `formatCellValue` | 渲染器同源的单元格格式化器；返回 `FormattedValue`（`{ text, color? }`，该类型同样随包导出——见[数据模型](/zh/packages/excel-preview/api/02-model)）                                                                                      |
| `configureWasm`   | 资源自托管配置（见[资源与自托管](/zh/packages/excel-preview/guide/02-assets)）                                                                                                                                                          |
| `getWasmLoader`   | 本包的 `WasmLoader` 单例（来自打包进本包 `dist` 的引擎层）。就绪状态用 `isReady` / `supported` 读取，配置用 `getOptions()`（状态机本身是私有的）。`ensureLoaded()` 也是公开的——可在首次预览前调用预热引擎，把 WASM 加载从首次解析中拿掉 |

`configureWasm` / `getWasmLoader` 来自打包进本包的引擎层：各 @marcusok 业务包各持一份副本与各自的 loader 实例——用哪个包就配置哪个。
