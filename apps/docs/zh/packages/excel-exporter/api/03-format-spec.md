# API：FormatSpec 值格式化

结构化、可跨线程的值格式化描述。Worker/Stream 路径建议使用 FormatSpec；函数形式的实际可用范围见文末「函数形式」。

## 类型定义

```ts
type FormatSpec =
  | { type: "enum"; map: Record<string, string>; fallback?: string }
  | { type: "date"; pattern?: string } // 默认 "yyyy-MM-dd"
  | { type: "datetime"; pattern?: string } // 默认 "yyyy-MM-dd HH:mm"
  | { type: "number"; decimals?: number; thousands?: boolean }
  | { type: "padding"; fill: string; length: number; align?: "left" | "right" };
```

## 各类型说明

### enum

```ts
{ type: "enum", map: { paid: "已支付", pending: "待支付" }, fallback: "未知" }
```

命中 `map` 输出映射值；未命中输出 `fallback`，无 fallback 时原样输出。缺失值——`null`/`undefined` 与空串/纯空白串——渲染为**空单元格**：既不查 `map` 也不落 `fallback`（`fallback` 只覆盖「有值但未映射」，不覆盖数据空洞，导出绝不会用有业务含义的假文案去填空洞）。

### date / datetime

```ts
{ type: "date" }                       // 默认 yyyy-MM-dd
{ type: "datetime", pattern: "yyyy-MM-dd HH:mm:ss" }
```

接受 `Date` / 可解析字符串 / 时间戳。两条路径都写入 Excel 日期序列，并把 pattern 作为 numFmt 附加——Workbook 路径通过单元格样式，Stream 路径通过最小 `styles.xml`——因此存储值与显示日期跨路径一致（受下文显式 `style.numFormat` 例外约束）。统一按 **UTC 分量**解释（与 Workbook 序列的 `dateToSerial` 口径一致，跨路径/跨时区显示相同）；ISO 日期字符串按 ECMA-262 解析为 UTC 午夜，详见[值格式化的时区约定](/zh/packages/excel-exporter/guide/04-formatting)。pattern 在两条路径上都作为数字格式交给 Excel，完整格式码语义（含引号字面量）均适用，不存在收窄的 token 子集。只有独立的 `formatDateByPattern` 辅助函数（导出给自定义消费方，`displayValue` 也使用它，但 stream 导出不使用）保留旧的小写化、有限 token 扫描行为——见其 API 条目。

同一区域还有一个边界：无法转换为日期的值（bigint、布尔、普通对象、Invalid Date——只接受 `Date` / 可解析字符串 / 时间戳）会以其字符串形态写入，而不会让导出失败。

### number

```ts
{ type: "number", decimals: 2, thousands: true }
```

`decimals` 默认 0，`thousands` 默认 false；`decimals` 校验为 0–100 的整数。两条路径都保留完整精度的存储值，并通过 Excel numFmt 渲染 `decimals`/`thousands`——Workbook 路径走单元格样式，Stream 路径走最小 `styles.xml`。因此跨 50,000 行不再舍入存储值，也不再丢失千分位。

日期/时间同理：两条路径都存 Excel 序列号（UTC 分量），pattern 作为 numFmt 附加，单元格类型与显示跨路径一致。

**显式 `CellStyle.numFormat` 与 FormatSpec 的优先级。** Workbook 路径上，列的显式 `style.numFormat` 优先，FormatSpec 推导出的格式码不会被注入；Stream 路径则完全丢弃 `CellStyle.numFormat`，使用 FormatSpec 的 pattern/decimals/thousands。两者同时设置且不一致时，存储值仍然一致，但跨 50k 阈值的显示会不同；请保持两者一致，或在 FormatSpec 列上省略 `style.numFormat`。

`null`/`undefined`（以及空串/纯空白串——数据库、表单、CSV 导入中最常见的缺失值形态）在所有路径下均渲染为空单元格——绝不会是 `0`。`padding` 对它们同样处理：空单元格，而不是 `"00000"` 这类填充出的假编号。`enum` 亦然：空单元格，绝不会是 `fallback`（见上文）。

### padding

```ts
{ type: "padding", fill: "0", length: 6 } // "42" -> "000042"
```

省略 `align`（默认）对应 `padStart`，在**左侧**补字符（值右对齐），适合工号、订单号的前导零场景；`align: "left"` 对应 `padEnd`，在**右侧**补字符（值左对齐）。补位从不截断：原值长度已达到或超过 `length` 时原样返回。`length` 校验为 0–10000 的整数，越界会前置报错，而不是等到渲染期才出错或静默不填充。

## 前置校验

整个 spec 家族在任何路由执行前统一校验——违规导出以 `{ success: false }` 失败并给出定位到列的报错，而不是拖到后面变成晦涩错误或静默降级的文件。除上文 `decimals` / `length` 的范围外还包括：

- `type` 必须是五个字面量之一——`"percent"` 这类拼写错误会被前置拒绝（否则值会被静默字符串化，数字列整列变文本且零告警）；
- `enum.map` 必须是非数组、非 null 的对象（缺失 / `null` / 数组都会被拒）；
- `date` / `datetime` 提供的 `pattern` 必须是字符串；
- `padding.fill` 必须是字符串——类型上本就必填，`undefined` 一并拦截（非字符串虽不抛错但会产出怪输出，故同样前置拦截）。

## 函数形式

```ts
format: (value, row) => string | number | boolean;
```

可以访问整行数据做条件格式化。函数无法穿过结构化克隆，因此各路径行为不同：

- **main 路径**（auto 模式下浏览器 < 20,000 行 / Node < 50,000 行；或任意行数显式 `mode: "main"`）：函数正常执行；
- **Node 的 stream 路径**（auto ≥ 50,000 行，或显式 `mode: "stream"`）：同样在主线程执行，函数正常执行；
- **worker 失败后的主线程重试**：函数正常执行——原始 options 里的函数仍在，被剥离的只是发给 worker 的副本；
- **浏览器 worker 路径**（auto ≥ 20,000 行，或显式 `mode: "worker"` / `mode: "stream"`）：函数会被**剥离并打印 `console.warn`**，该列以原始值导出（不会报错，也不会回落到 main）。

需要 worker 路径保留格式时，请改写为 FormatSpec。

函数抛错**不会**被隔离：异常会从构建过程抛出，并可能触发正常的降级/失败链，因此函数应自行兜底、处理坏值。返回 `null` / `undefined`（或空串）会渲染为空单元格。
