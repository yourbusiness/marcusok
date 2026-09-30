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

| 选项            | 类型                                        | 默认值           | 说明                                                                                                                                                                                                                           |
| --------------- | ------------------------------------------- | ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `source`        | `File \| Blob \| Uint8Array \| ArrayBuffer` | —（必填）        | 文件字节                                                                                                                                                                                                                       |
| `password`      | `string`                                    | —                | 加密工作簿的密码（Agile AES-256）                                                                                                                                                                                              |
| `sheet`         | `string \| number`                          | 文件 `activeTab` | 初始 sheet（名称或 0 起索引）。名称不存在或索引越界时静默回退文件 `activeTab`——与实例方法 `setSheet()` 不同，后者会把无效入参经 `onError` 报出                                                                                 |
| `showHeaders`   | `boolean`                                   | `true`           | 行列表头（A/B/C + 1/2/3）                                                                                                                                                                                                      |
| `showGridLines` | `boolean`                                   | 遵循文件         | 网格线                                                                                                                                                                                                                         |
| `showTabs`      | `boolean`                                   | `true`           | sheet 页签栏（隐藏表永不出现）                                                                                                                                                                                                 |
| `onParsed`      | `(info: PreviewParsedInfo) => void`         | —                | 渲染就绪回调：首次解析渲染完成后与**每次 sheet 切换完成后**都会触发（`duration.parse` 复用首次解析耗时，`duration.total` 则一直从首次加载起累计——要衡量本次切换本身的耗时请用 `duration.render`）；上报 sheet 列表、规模与耗时 |
| `onError`       | `(error: PreviewError) => void`             | —                | 失败回调（错误码见下）                                                                                                                                                                                                         |

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

| 错误码               | 含义                                                                                                                  |
| -------------------- | --------------------------------------------------------------------------------------------------------------------- |
| `PASSWORD_PROTECTED` | 加密工作簿，`password` 缺失或错误                                                                                     |
| `LEGACY_FORMAT`      | 旧版 `.xls`（BIFF8）——请另存为 `.xlsx`                                                                                |
| `CORRUPT`            | ZIP 但不是有效 xlsx：部件缺失/损坏（`.ods` / `.docx` 同形包、截断文件），或工作簿一个 sheet 都没有                    |
| `UNSUPPORTED`        | 既不是 ZIP 也不是可识别的纯文本——含 XML/HTML 伪表格（SpreadsheetML 2003、HTML 表格另存为 `.xls`）与 UTF-16 编码的 CSV |
| `WASM`               | 当前环境不支持 WebAssembly，或引擎加载失败（资产 404、CSP 禁止 WebAssembly、网络失败）                                |
| `UNKNOWN`            | 其他；底层错误信息原样透传（失败来自预览自身 boot 路径时，原始错误对象在 `.cause`）                                   |

## PreviewInstance

```ts
interface PreviewInstance {
  destroy(): void; // 卸载 DOM、释放资源
  setSheet(nameOrIndex: string | number): void; // 切换 sheet
  getSheetNames(): string[]; // 全部 sheet，文件顺序
}
```

`destroy()` 移除渲染 DOM。解析 worker 是模块级共享资源（与导出包一致，有意跨实例保温）；解析完成后不持有任何文件数据。

## parseWorkbookBytes

低层解析，不做 DOM：

```ts
import { parseWorkbookBytes } from "@marcusok/excel-preview";

const workbook = await parseWorkbookBytes(bytes, { password: "…" });
```

失败时抛出带 `.code` 属性的 `Error`（错误码同上）。模型字段见[数据模型](/zh/packages/excel-preview/api/02-model)。

## 其他导出

| 导出              | 说明                                                                                                                                                      |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `formatCellValue` | 渲染器同源的单元格格式化器（见[数据模型](/zh/packages/excel-preview/api/02-model)）                                                                       |
| `configureWasm`   | 资源自托管配置（见[资源与自托管](/zh/packages/excel-preview/guide/02-assets)）                                                                            |
| `getWasmLoader`   | 共享的 `WasmLoader` 单例——与 `@marcusok/xlsx-core` 导出的是同一对象。就绪状态用 `isReady` / `supported` 读取，配置用 `getOptions()`（状态机本身是私有的） |

`configureWasm` / `getWasmLoader` 是 `@marcusok/xlsx-core` 的再导出：从任一包导入，配置的都是同一个共享 loader。
