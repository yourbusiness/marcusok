# 数据模型（PreviewWorkbook）

`parseWorkbookBytes`（与 Worker）产出纯 JSON 模型——结构化克隆安全、框架无关，与内置渲染器消费的数据完全一致。本页是字段级参考：逐字段说明解析层实际产出的取值，以及内置渲染器对它的处理方式，因此仅凭该模型就能写出自研渲染器。

## 结构

```ts
interface PreviewWorkbook {
  sheets: PreviewSheet[];
  activeSheetIndex: number; // 0 起，来自文件 activeTab（越界已钳制到 0）
  dateSystem: "date1900" | "date1904"; // 决定序列号 → 日期的换算
}

interface PreviewSheet {
  name: string;
  visible: boolean; // hidden / veryHidden 保留在列表里
  showGridLines: boolean; // 取自文件视图设置；缺省 true
  rightToLeft: boolean;
  rowCount: number; // 1 起数据边界——见下文「边界」
  colCount: number;
  rows: PreviewRow[]; // 稀疏——见下文「单元格」
  colSpans: PreviewColSpan[]; // 对应文件 <col>，min/max 1 起
  merges: PreviewMerge[]; // 0 起 { row, col, rowSpan, colSpan }
  frozenRows: number; // 0 = 无
  frozenCols: number;
  styles: PreviewStyles; // 工作簿级：各 sheet 共享
}

interface PreviewRow {
  index: number; // 1 起
  height: number | null; // pt；null → 默认（15pt）
  hidden: boolean;
  cells: PreviewCell[]; // 按列稀疏，已按 col 升序
}

interface PreviewCell {
  col: number; // 0 起
  type: PreviewCellType; // "number" | "string" | "boolean" | "error" | "formulaStr"
  value: string | null; // 原始字符串形态；数字保持全精度
  styleIndex: number | null; // styles.xfs 下标；null = 默认样式
}

interface PreviewColSpan {
  min: number; // 1 起，闭区间
  max: number; // 1 起，闭区间
  width: number; // Excel 字符宽度单位；文件未写时为 8.43
  hidden: boolean; // 隐藏列同样在模型里（渲染宽度为 0）
  customWidth: boolean; // 仅元信息——见「列跨度」
}

interface PreviewMerge {
  row: number; // 0 起，主格（左上角）
  col: number;
  rowSpan: number; // >= 1
  colSpan: number;
}

interface PreviewStyles {
  fonts: PreviewFont[];
  fills: PreviewFill[];
  borders: PreviewBorder[];
  xfs: PreviewXf[]; // 单元格样式，由 PreviewCell.styleIndex 索引
}
```

## 边界（rowCount / colCount）

两者都是 **1 起的「数据边界」**，不是数组长度——而且允许明显大于 `rows` 数组所示的范围：

| 边界       | 取以下各项的最大值                                                                      |
| ---------- | --------------------------------------------------------------------------------------- |
| `rowCount` | 现有各 `PreviewRow.index` 的最大值，以及每个合并的 `row + rowSpan`                      |
| `colCount` | 每个单元格的 `col + 1`、每个跨度的 `PreviewColSpan.max`，以及每个合并的 `col + colSpan` |

有两个渲染方不能踩空的推论：

- **合并区会撑大边界。** 只有 A1 有值、且存在 A1:C3 合并的 sheet，会报告 `rowCount: 3, colCount: 3`——尽管 `rows` 只有一项、第 2/3 列上根本不存在单元格。Excel 对纯合并不写覆盖格，边界是合并块范围的唯一记录；布局若不以它为准，合并块会被截成 1×1。
- **`<col>` 跨度会撑大 `colCount`。** 声明了 `<col min="1" max="20">` 的文件会报告 `colCount >= 20`，与这些列里有没有内容无关。

真正空白的 sheet（无行、无跨度、无合并）可以是 `0`。内置布局还会把边界下限压到冻结行/列数，以及 `1`：冻结窗格声明的冻结行数可能超过 sheet 的实际内容。

## 单元格

**`type` 是归一化后的取值，不是引擎的 `cellType`。** 解析层把引擎各种「出字符串」的形态合并为一种：

| 引擎 `cellType`                                        | 模型 `type`    |
| ------------------------------------------------------ | -------------- |
| `number`                                               | `"number"`     |
| `boolean`                                              | `"boolean"`    |
| `error`                                                | `"error"`      |
| `formulaStr`                                           | `"formulaStr"` |
| `sharedString` / `inlineStr` / `stub` / 其它一切未知值 | `"string"`     |

也就是说模型不再记录字符串的来源——共享字符串表里的条目与内联字符串都是 `"string"`；引擎返回未知类型时也降级为 `"string"`，而不是报错。注意 `"stub"`（被引用但为空的共享字符串）同样并入 `"string"`；这类格子通常在更早一步就消失了——只有「有值」**或**「有样式」的单元格才会进模型。

**`value` 恒为原始字符串形态。** 数字保持全精度（`"45678.5"`），布尔为引擎原始形态（多数文件里是 `"1"` / `"0"`），`error` 携带错误字面量（`"#DIV/0!"`），`formulaStr` 携带公式的字符串缓存值。数值缓存的公式单元格以 `type: "number"` 携带该缓存值出现——与 Excel 显示一致。

**`value: null` 表示「有样式但无值」**，不是「缺席」：无值但有样式的单元格仍会进模型，以便渲染其填充与边框（格式化结果为空串）。既无值又无样式的单元格被整体丢弃。

**`styleIndex` 是 `styles.xfs` 的下标；`null` 表示默认样式。** 对于 `null` 不要退而索引 `xfs[0]`——见下文「自行格式化单元格」。

`type` 同时决定内置渲染器 General 对齐的分流：数字右对齐，`string` 与 `formulaStr` 左对齐，`boolean` 与 `error` 居中。仅在 xf 未携带显式水平对齐时适用。

## 列跨度

`colSpans` 与文件的 `<col>` 元素一一对应，按文档序排列——**不会**按列展开，因此一项可能覆盖上百列。

- `min` / `max` 为 1 起闭区间：`{ min: 2, max: 5 }` 覆盖 B–E 列（0 起的 1–4）。
- `width` 用 Excel 字符宽度单位，即文件自身 `width` 属性的单位。文件未写时解析层兜底 `8.43`——恰好是 Excel 的默认列宽，因此该兜底在视觉上是空操作。
- `hidden` 列保留自己的位置，宽度为 **0**（不从坐标空间中移除；见下文「写自定义渲染器」）。
- `customWidth` 仅作**元信息**保留。不要用它把关列宽：实测引擎对两种形态都返回 `customWidth: false`，用它把关会静默丢掉第三方产物「只写 width 不写标志」的列宽。只要 `<col>` 出现，就采用它的宽度。

按内置渲染器的口径换算像素：`round(width * 7 + 5)`（Calibri 11 @ 96dpi，最大数字宽 7）。列级样式（`<col style="…">`）不在模型内。

## 样式

`PreviewStyles` 是工作簿级且共享的：文件只有一份 `styles.xml`，各 sheet 指向同一对象。不要按 sheet 深拷贝——样式表只编译一次。

### 单元格样式（PreviewXf）

| 字段         | 类型                    | 含义                                                     |
| ------------ | ----------------------- | -------------------------------------------------------- |
| `fontId`     | `number`                | `styles.fonts` 的下标                                    |
| `fillId`     | `number`                | `styles.fills` 的下标                                    |
| `borderId`   | `number`                | `styles.borders` 的下标                                  |
| `numFmtCode` | `string`                | 已解析的格式码；永不为 `undefined`（下限是 `"General"`） |
| `alignment`  | `AlignmentData \| null` | 显式对齐，`null` 表示未设置                              |

id 字段直接来自文件的 `cellXfs`（属性缺失时兜底 `0`），而三个集合按 `styles.xml` 文档序重建，以保证 id 对齐——自闭合的默认项也会作为空项计入。即便如此也要防御性索引：越界的 id 取到 `undefined`，内置编译器对空槽位不输出任何声明。

### 字体（PreviewFont）

| 字段                                       | 类型             | 含义                                                        |
| ------------------------------------------ | ---------------- | ----------------------------------------------------------- |
| `name`                                     | `string \| null` | 字体名；`null` → 用渲染器自带的字体栈（Calibri、Segoe UI…） |
| `size`                                     | `number \| null` | 字号，单位 **pt**；**`null` = 默认 11pt**                   |
| `bold` / `italic` / `underline` / `strike` | `boolean`        | 每项一个标志；下划线与删除线会合并为一条 `text-decoration`  |
| `color`                                    | `string \| null` | 已解析的 CSS 色；**`null` = 自动 → 渲染为黑**               |

恶意文件的字体名在拼进 CSS `font-family` 声明前会被剥掉 `"` 与 `\`（未转义的引号会提前终止生成的规则，让文件得以注入任意声明）。其余插值要么是校验过的十六进制色，要么是固定关键字。

### 填充（PreviewFill）

以 `kind` 为判别式的联合类型——四个变体的字段结构各不相同，必须先收窄：

```ts
type PreviewFill =
  | { kind: "none" }
  | { kind: "solid"; fgColor: string | null }
  | { kind: "pattern"; fgColor: string | null; bgColor: string | null }
  | {
      kind: "gradient";
      degree: number;
      stops: { position: number; color: string }[];
    };
```

| `kind`     | 产出条件                           | 内置渲染器的画法                                                                         |
| ---------- | ---------------------------------- | ---------------------------------------------------------------------------------------- |
| `none`     | 无填充，或图案类型缺失 / 为 `none` | 不画                                                                                     |
| `solid`    | 图案类型为 `solid`                 | `background-color: fgColor`；`fgColor` 为 `null`（自动色）时按**无填充**处理，而不是涂黑 |
| `pattern`  | 其它任何图案类型                   | 按纯色近似：`background-color: bgColor ?? fgColor ?? #ffffff`——斜纹本身不还原            |
| `gradient` | 带至少一个 stop 的 `gradientFill`  | `linear-gradient((degree + 90) % 360deg, …)`，stop 位置换算成百分比                      |

自研时值得知道的渐变细节：

- `position` 取值 `0..1`。解析层对缺失的 position 按均分补齐（`i / (n - 1)`），并按 position 排序；颜色无法解析的 stop 兜底 `#000000`。
- `degree` 依 ECMA-376：`0` = 左→右、`90` = 上→下。CSS 的 `0deg` 朝上、`90deg` 朝右，故有 `+ 90` 的平移。
- 渐变填充与对角线边框都必须落在 `background-image`。两条独立声明不会叠加——后写的胜出——因此必须合成一条逗号分隔的多层声明（对角线在上层）。

### 边框（PreviewBorder）

五条边，每条要么是 `{ style: BorderStyle; color: string | null }`，要么是 `null`：

| 边                                  | 含义                                          |
| ----------------------------------- | --------------------------------------------- |
| `left` / `right` / `top` / `bottom` | 格子矩形的某条边；`null` = 该边无边框         |
| `diagonal`                          | 格子的对角线；`null` = 无（方向标志规则见下） |

文件未声明该边样式、或声明为 `style="none"` 时，该边为 `null`。`color: null` 是自动色，渲染为 `#000000`。

四条边映射到 CSS 的 `border-<side>` 声明：

| Excel `style`                                                       | 内置渲染器输出的 CSS |
| ------------------------------------------------------------------- | -------------------- |
| `thin`、`hair`                                                      | `1px solid`          |
| `medium`                                                            | `2px solid`          |
| `thick`                                                             | `3px solid`          |
| `double`                                                            | `3px double`         |
| `dashed`、`dashDot`、`dashDotDot`                                   | `1px dashed`         |
| `mediumDashDot`、`mediumDashDotDot`、`mediumDashed`、`slantDashDot` | `2px dashed`         |
| 其它无法识别的取值                                                  | `1px solid`          |

CSS 没有原生对角线，因此 `diagonal` 不是边框，而是横跨格子矩形的一层细线性渐变（`medium` / `thick` 为 2px，其余 1px）。方向由另一个可选字段决定：

- **`diagonalUp` 是方向标志：`true` = `/`（Excel 的 `diagonalUp`）；缺省或 `false` = `\`（Excel 的 `diagonalDown`，更常见的形态）。**
- **只声明了对角线样式、却既没有 up 也没有 down 标志的文件，完全不显示对角线**——Excel 本身就不画。解析层已在该情形把 `diagonal` 置为 `null`，渲染期无需再自行判断；也正因如此，`diagonalUp` 只在 `diagonal` 非 null 时才有意义。
- 两个标志同置（X 形）时单条渐变表达不了两条线，取 down 形态，结果为 `\`。
- 标志规则有一个例外：`styles.xml` 完全读不出来时，解析层回落到引擎的边框数组，而该数组不携带 up/down 标志。这条降级路径保留引擎给出的对角线并固定 `diagonalUp: false`（`\`）——两种取法都是猜测，且只影响样式部件不可读的文件。

### 对齐（PreviewXf.alignment）

类型是 `AlignmentData | null`，从 `@marcusok/xlsx-core` 原样再导出（解析层未另行声明）。`null` 表示文件未设置对齐，此时按渲染器「按类型分流」的 General 规则处理。引擎类型声明了六个字段，但**渲染层只消费其中五个**：

| 字段           | 类型                                                                                                               | 是否消费 | 渲染层行为                                                                                                                                                                                                                  |
| -------------- | ------------------------------------------------------------------------------------------------------------------ | -------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `horizontal`   | `"general" \| "left" \| "center" \| "right" \| "fill" \| "justify" \| "centerContinuous" \| "distributed" \| null` | 是       | `left` / `center` / `right` 落在 flex 主轴上；`justify` 与 `distributed` 保持左对齐并追加 `text-align: justify`；`general`、`fill`、`centerContinuous` 落回按类型分流的规则（`fill` 与 `centerContinuous` 按 General 近似） |
| `vertical`     | `"top" \| "center" \| "bottom" \| "justify" \| "distributed" \| null`                                              | 是       | `top` / `center` / `bottom` 落在 flex 交叉轴上；`justify` 与 `distributed` 无 CSS 等价物，按 `center` 近似                                                                                                                  |
| `wrapText`     | `boolean`                                                                                                          | 是       | 追加 `white-space: pre-wrap` + `word-break: break-word`                                                                                                                                                                     |
| `indent`       | `number \| null`                                                                                                   | 是       | 见下                                                                                                                                                                                                                        |
| `textRotation` | `number \| null`                                                                                                   | 是       | 见下                                                                                                                                                                                                                        |
| `shrinkToFit`  | `boolean`                                                                                                          | **否**   | 引擎类型声明了，但渲染器忽略：不会缩放文本去适应列宽，溢出仍走常规的裁剪/溢出规则                                                                                                                                           |

`indent` 的折算口径：Excel 一个缩进单位约等于三个空格字符宽，故每级转为 `padding-inline-start: indent * 3ch`（`ch` 相对单元格自身字号，正是它适合做代理的原因）。`horizontal` 为 `"right"` 时改为在行尾留白——Excel 的右对齐文本从右侧缩进。

`textRotation` 的映射：

| 取值        | 渲染层行为                                                                      |
| ----------- | ------------------------------------------------------------------------------- |
| `null`、`0` | 不旋转                                                                          |
| `1`–`90`    | 逆时针：`rotate(-n deg)`，原点在左下，文本保持单行                              |
| `91`–`180`  | 顺时针：`rotate((n - 90) deg)`，原点在右下，文本保持单行                        |
| `255`       | 竖排堆叠文本（`writing-mode: vertical-rl`）——Excel 的「竖排文字」选项，不是旋转 |

旋转无法反映到格子矩形上，因此旋转文本保持单行、放不下时按格子矩形裁剪——与 Excel 在矩形对角线附近的差异见[能力边界](/zh/packages/excel-preview/guide/04-limits)。

## 消费方要点

- **值是原始字符串。** 数字保持全精度字符串形态；需要数值时自行 `Number(value)`。渲染器的格式化路径就是这么做的。
- **`numFmtCode` 已预解析。** 内置 id 已过 Excel 行为表（14 → `m/d/yyyy`），自定义 id 来自文件 `numFmts`。永远不会是 `undefined`——下限是 `"General"`。
- **颜色是 CSS 字符串**（`#rrggbb` / `rgba(...)`），已经过主题/indexed 覆盖层解析；`null` 表示自动（渲染默认：黑字、无填充）。
- **`formulaStr` 单元格**是公式的字符串缓存值；数字缓存的公式单元格以 `type: "number"` 携带缓存值出现（与 Excel 显示一致）。
- **空单元格缺席**——模型两个维度都是稀疏的；渲染方应把「缺席」当作默认样式的空格。
- **模型是纯 JSON**，没有类实例、也没有 `Date` 对象，因此可以原样 `structuredClone`、缓存或在上下文之间传递。它同时按只读约定使用：请自行派生索引结构，不要在上面挂载数据。

## 自行格式化单元格

渲染器使用的格式化器原样导出，供自研渲染器复用：

```ts
import { formatCellValue } from "@marcusok/excel-preview";

// 与渲染器口径一致：没有 xf 就没有格式码，按 "General" 处理——不要退而取
// xfs[0]，它只是恰好在多数文件里是 Normal 样式。
const xf =
  cell.styleIndex != null ? sheet.styles.xfs[cell.styleIndex] : undefined;

const { text, color } = formatCellValue(
  cell.type,
  cell.value,
  xf?.numFmtCode ?? "General",
  workbook.dateSystem,
);
// text: 显示文本；color: 格式含 [Red] 类颜色段时的 "#ff0000"
```

| 参数         | 类型                       | 说明                                                                                            |
| ------------ | -------------------------- | ----------------------------------------------------------------------------------------------- |
| `type`       | `PreviewCellType`          | 直接传 `cell.type`：格式化器按它分支（布尔 → `TRUE`/`FALSE`、错误原样、General 数字走数值路径） |
| `value`      | `string \| null`           | 直接传 `cell.value`；`null` 或 `""` 返回 `{ text: "" }`                                         |
| `numFmtCode` | `string`                   | 该格 xf 的已解析格式码；空串或缺失按 `"General"` 处理，且不会因此抛错                           |
| `dateSystem` | `"date1900" \| "date1904"` | 传 `workbook.dateSystem`；1904 系统的序列号在日期格式化前先 +1462 天平移                        |

返回 `FormattedValue`（该类型已从包内导出）：

```ts
interface FormattedValue {
  text: string; // 显示文本
  color?: string; // 格式含 [Red] / [Color 3] 颜色段时的 CSS 色
}
```

`color` 是可选的，仅在选中的格式段携带颜色括号时给出。请把它作为内联颜色覆盖 xf 的字体色（渲染器写的是 `el.style.color`），并且只在有值时才写——用 `undefined` 覆盖会把合法的黑色也一并抹掉。

格式化器恢复了哪些、近似了哪些，完整口径见[格式保真](/zh/packages/excel-preview/guide/03-format-fidelity)。

## 写自定义渲染器

如果你要基于该模型自建 DOM，以下性质最容易踩坑：

1. **模型两个方向都是稀疏的。** `rows` 只包含有内容、有自定义行高或有隐藏标志的行；`cells` 只包含有值或有样式的列。遍历要用 `index` / `col` 字段，绝不能用数组下标——行已按 `index` 排序，但空档是真实存在的。
2. **「缺席」与「有样式但无值」不是一回事。** 缺失的 `(row, col)` 是默认样式的空格：画网格底色与网格线，而不是一个默认样式的格子。而模型里 `value: null` 的格子是「有格式的空格」，它的填充与边框必须照画。
3. **撑起网格尺寸的是 `rowCount` / `colCount`，不是数组。** 它们已并入合并区与 `<col>` 跨度的范围（见上文「边界」），是可滚动范围的唯一正确来源。
4. **数字是字符串。** 用 `Number(value)` 转换；不要反向解析显示文本。保留原串正是为了不丢失 `Number` 往返可能带来的精度损失。
5. **颜色已是解析好的 CSS 字符串。** 原样使用——显式 RGB 是 `#rrggbb`，带 tint 的主题色是 `rgba(...)`。`null` 表示自动：字体与边框为黑，填充为不画。
6. **`styleIndex: null` 表示默认样式**，不是 `xfs[0]`。跳过对应的类名与 xf 查找，并按 `"General"` 格式化。
7. **样式按工作簿编译一次，不要逐格编译。** 内置渲染器为每个 xf 下标生成一个 CSS 类，单元格只挂类名；逐格内联样式对象会在每次虚拟滚动重建时重算，而类名查找是免费的、且能被浏览器缓存计算样式。只有格式码颜色（来自 `FormattedValue.color`）需要内联。
8. **跳过被合并覆盖的格子。** 一个合并只产出主格一条记录，被覆盖的坐标在 `rows` 里不存在。请按合并后的尺寸绘制主格（跨过列宽与行高之和），让它自然盖住内部——包括左上角落在视口外的那些主格，否则跨越视口的合并块会画不全。
9. **隐藏行列保留自己的槽位，尺寸为 0。** 它们仍占据坐标空间中的下标；把它们从前缀和里剔掉会让其后所有行列整体位移。`<col hidden>` 是宽度为 0 的列，不是不存在的列。
10. **把模型当作只读。** 它经结构化克隆从 worker 传来，且可能被多次渲染复用；请自建查找结构（例如以 `row:col` 为键的 `Map`），而不是就地修改。
