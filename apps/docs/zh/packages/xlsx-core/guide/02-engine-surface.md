# 引擎再导出面

下面全部是 [modern-xlsx](https://github.com/ABCrimson/modern-xlsx) 公开 API 的**再导出**——即各 @marcusok 包共用的那部分，加上它们需要的类型。此处给出的签名就是引擎自身的签名；本包没有加任何包装层，因此引擎文档说明的行为原样适用。

## 工作簿、读取与写入

```ts
import {
  Workbook,
  readBuffer,
  initWasm,
  initWasmSync,
} from "@marcusok/xlsx-core";

function readBuffer(data: Uint8Array, options?: ReadOptions): Promise<Workbook>;
class Workbook {
  constructor(data?: Partial<WorkbookData>);
  readonly sheetNames: readonly string[];
  getSheet(name: string): Worksheet | undefined;
  addSheet(name: string): Worksheet;
  toBuffer(options?: WriteOptions): Promise<Uint8Array>;
}
function initWasm(wasmSource?: string | URL | Response): Promise<void>;
function initWasmSync(module: WebAssembly.Module | BufferSource): void;
```

```ts
const workbook = await readBuffer(bytes); // ReadOptions = { password?: string }
const sheet = workbook.getSheet("Sheet1") ?? workbook.addSheet("Sheet1");
sheet.cell("A1").value = "Hello";
const out = await workbook.toBuffer(); // WriteOptions = { password?: string }
```

`initWasm` / `initWasmSync` 是引擎自己的初始化调用。通常用不到它们：`getWasmLoader().ensureLoaded()` 会替你调用（见 [WASM 加载器指南](/zh/packages/xlsx-core/guide/01-loader)）。例外是**单文件自包含的 worker 入口**——它无法使用 loader 的 Node 自动初始化路径，因此会带显式 URL 调用 `initWasm(wasmUrl)`，`@marcusok/excel-preview` 的解析 worker 正是如此。`initWasmSync` 则是 loader 的 Node 自动初始化内部调用的那个。

## 单元格与区域引用

```ts
import {
  encodeCellRef,
  decodeCellRef,
  encodeRange,
  decodeRange,
  columnToLetter,
  letterToColumn,
} from "@marcusok/xlsx-core";

function encodeCellRef(row: number, col: number): string; // (0, 0) → "A1"
function decodeCellRef(ref: string): CellAddress; // "A1" → { row: 0, col: 0 }
function encodeRange(start: CellAddress, end: CellAddress): string; // → "A1:C10"
function decodeRange(range: string): CellRange; // → { start: CellAddress, end: CellAddress }
function columnToLetter(col: number): string; // 0 → "A", 26 → "AA"
function letterToColumn(letter: string): number; // "A" → 0, "AA" → 26
```

所有下标都是 **0 基**；A1 字符串是 1 基的，这正是 `encodeCellRef(0, 0)` 等于 `"A1"` 的原因。`CellAddress` 为 `{ readonly row: number; readonly col: number }`，`CellRange` 为 `{ readonly start: CellAddress; readonly end: CellAddress }`。

这一组是导出包用得最重的：它用 `encodeCellRef` 生成写入单元格的 A1 引用，用 `decodeCellRef` 解析调用方要求写入的合并区域。`columnToLetter` / `letterToColumn` 目前两个业务包都没用到——它们是为自绘列头的渲染器准备的。

## 数字格式化

```ts
import {
  formatCell,
  formatCellRich,
  getBuiltinFormat,
  loadFormatTable,
  isDateFormatId,
  isDateFormatCode,
} from "@marcusok/xlsx-core";

function formatCell(
  value: string | number | boolean | null,
  format: string | number, // 格式码，或内置格式 id
  opts?: FormatCellOptions, // { dateSystem?: DateSystem }
): string;

function formatCellRich(
  value: string | number | boolean | null,
  format: string | number,
  opts?: FormatCellOptions,
): FormatCellResult; // { text: string; color?: string }

function getBuiltinFormat(id: number): string | undefined;
function loadFormatTable(table: Record<number, string>): void;
function isDateFormatId(numFmtId: number): boolean;
function isDateFormatCode(formatCode: string): boolean;
```

`formatCellRich` 额外支持条件段（`[>100]#,##0;[<=100]0.00`）与方括号颜色指令（`[Red]`、`[Color3]`），把颜色名连同文本一起返回——预览的格式化层就是建立在它之上的：

```ts
const { text, color } = formatCellRich(1234.5, "#,##0.00");
// text: "1,234.50"
```

以下注意事项属于引擎而非本包：`formatCell` / `formatCellRich` 遇到 `null` / `undefined` 的格式码会抛 `TypeError`（务必先把 id 解析成格式串），且多段格式的数字段内字面量会被引擎丢弃。预览正是针对这些缺陷做了补偿层——在依赖 `formatCellRich` 的输出以求与 Excel 显示一致之前，请先读[格式忠实度](/zh/packages/excel-preview/guide/03-format-fidelity)。

`isDateFormatId` 只查内置 id；自定义格式码请用 `isDateFormatCode`，它会扫描格式串里的日期/时间记号并跳过引号串。这两个目前都没有被两个业务包使用——它们进入导出面，是因为任何"这个数值是不是日期"的判断都需要它们。

## 日期序列号

```ts
import { serialToDate, dateToSerial } from "@marcusok/xlsx-core";

function dateToSerial(date: Date | TemporalLike, system?: DateSystem): number;
function serialToDate(serial: number, system?: DateSystem): Date; // UTC
```

```ts
dateToSerial(new Date(Date.UTC(2024, 0, 1))); // 45292
serialToDate(45292); // 2024-01-01T00:00:00.000Z
```

`DateSystem` 为 `"date1900" | "date1904"`，默认 `date1900`（Windows Excel）。两个函数都接受显式传入；两套体系相差 1462 天，且 1900 体系还带着 Excel 那个虚构的闰日。`dateToSerial` 另外接受鸭子类型化的 `Temporal.PlainDate` / `PlainDateTime`。导出包用 `dateToSerial` 把 `Date` 单元格值转成序列号；预览在显示侧用 `serialToDate`。

## sheet 助手

```ts
import { sheetAddAoa } from "@marcusok/xlsx-core";

function sheetAddAoa(
  ws: Worksheet,
  data: unknown[][],
  opts?: SheetAddAoaOptions, // { origin?: string }
): void;
```

```ts
const ws = new Workbook().addSheet("Data");
sheetAddAoa(ws, [
  ["Region", "Amount"],
  ["EMEA", 1200],
]); // 默认从 "A1" 开始写入
sheetAddAoa(ws, [["APAC", 800]], { origin: "A4" });
```

不传 `origin` 时，数据追加在既有内容之后。这是导出包构建器的基础原语：它按 sheet 把数组的数组交给引擎，而不是逐格写入。

## 类型化错误

```ts
import {
  ModernXlsxError,
  LEGACY_FORMAT,
  UNRECOGNIZED_FORMAT,
  PASSWORD_PROTECTED,
  WASM_INIT_FAILED,
} from "@marcusok/xlsx-core";

class ModernXlsxError extends Error {
  readonly code: string;
  constructor(code: string, message: string);
  static fromWasmError(err: unknown): ModernXlsxError;
}

LEGACY_FORMAT; // "LEGACY_FORMAT"——旧版 .xls（BIFF8），不支持
UNRECOGNIZED_FORMAT; // "UNRECOGNIZED_FORMAT"——既不是 ZIP 也不是 OLE2
PASSWORD_PROTECTED; // "PASSWORD_PROTECTED"——工作簿已加密
WASM_INIT_FAILED; // "WASM_INIT_FAILED"——WASM 未初始化 / 初始化失败
```

Rust 核心抛出的错误以 `"[CODE] message"` 形式到达；`ModernXlsxError.fromWasmError` 会把它解析成 `code` 加干净的 message，消息中识别不出码时回退为 `WASM_ERROR`。这四个常量是以**值**形式导出的码；引擎的错误码联合体更大（`ZIP_READ`、`ZIP_ENTRY`、`MISSING_PART`、`XML_PARSE`、`INVALID_*`……），但那些在引擎里只是类型，需要的消费方只能按字面字符串匹配——`@marcusok/excel-preview` 的错误归一模块把结构类失败映射为自己的 `CORRUPT` 码时就是这么做的（[预览错误码](/zh/packages/excel-preview/api/01-create-preview#previewerror-错误码)）。

一定要按 `code` 分流，不要按 message：消息随输入形态变化，码才是稳定契约。

## 类型再导出

与上面的运行时值一起导出，供不想直接依赖 `modern-xlsx` 的消费方给自己的代码加类型：

| 分组       | 类型                                                                                                                                            |
| ---------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| 读 / 写    | `ReadOptions`、`WriteOptions`、`WorksheetData`                                                                                                  |
| 单元格与表 | `Worksheet`、`RowData`、`CellData`、`CellType`、`ColumnInfo`、`HyperlinkData`、`SheetViewData`、`FrozenPane`                                    |
| 样式       | `StylesData`、`CellXfData`、`FontData`、`FillData`、`BorderData`、`BorderSideData`、`BorderStyle`、`AlignmentData`、`ThemeColorsData`、`NumFmt` |
| 格式化     | `FormatCellOptions`、`FormatCellResult`                                                                                                         |
| 日期       | `DateSystem`                                                                                                                                    |

```ts
import type { CellData, CellType, Worksheet } from "@marcusok/xlsx-core";
```

## 为什么从这里再导出，而不是直接依赖 modern-xlsx

- **一个版本，一处决策。** `modern-xlsx` 在本包内被钉死在精确版本上。转手发布它的 API，意味着引擎版本是一次性决定，而不是每个消费方各自声明、各自升级。
- **一份引擎实例。** 引擎运行时住在本包的 `dist/`，且消费方把 core 保持为 `external`，因此两个不同 @marcusok 包里的 `readBuffer` 解析到同一个模块——模块级的 WASM 初始化状态也与 loader 共用。
- **不与上游 engines 冲突。** `modern-xlsx` 声明 `engines.node >= 24`；由于它的运行时被打包进来（而该声明只能顺着本包钉死的依赖被看到，本包自己的 `engines` 是 `>= 22`），消费方不会被推向一个自己并不运行的 Node 版本。

代价是这层导出面是精选子集：引擎导出、但本仓库用不到的符号不会被再发布，所以取用前请先查本页。
