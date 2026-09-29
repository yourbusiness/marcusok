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

| 选项            | 类型                                        | 默认值           | 说明                              |
| --------------- | ------------------------------------------- | ---------------- | --------------------------------- |
| `source`        | `File \| Blob \| Uint8Array \| ArrayBuffer` | —（必填）        | 文件字节                          |
| `password`      | `string`                                    | —                | 加密工作簿的密码（Agile AES-256） |
| `sheet`         | `string \| number`                          | 文件 `activeTab` | 初始 sheet（名称或 0 起索引）     |
| `showHeaders`   | `boolean`                                   | `true`           | 行列表头（A/B/C + 1/2/3）         |
| `showGridLines` | `boolean`                                   | 遵循文件         | 网格线                            |
| `showTabs`      | `boolean`                                   | `true`           | sheet 页签栏（隐藏表永不出现）    |
| `onParsed`      | `(info: PreviewParsedInfo) => void`         | —                | 成功回调：sheet 列表、规模与耗时  |
| `onError`       | `(error: PreviewError) => void`             | —                | 失败回调（错误码见下）            |

`PreviewParsedInfo`：

```ts
interface PreviewParsedInfo {
  sheetNames: string[];
  sheetCount: number;
  rowCount: number; // 当前 sheet
  colCount: number;
  duration: { parse: number; render: number; total: number }; // 毫秒
}
```

## PreviewError 错误码

| 错误码               | 含义                                   |
| -------------------- | -------------------------------------- |
| `PASSWORD_PROTECTED` | 加密工作簿，`password` 缺失或错误      |
| `LEGACY_FORMAT`      | 旧版 `.xls`（BIFF8）——请另存为 `.xlsx` |
| `CORRUPT`            | 不是有效的 ZIP/xlsx 结构               |
| `UNSUPPORTED`        | 既不是 xlsx/zip 也不是纯文本（CSV）    |
| `WASM`               | WebAssembly 不可用（仅主线程回退路径） |
| `UNKNOWN`            | 其他（原始错误在 `.cause`，可得时）    |

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
