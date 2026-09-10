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

命中 `map` 输出映射值；未命中输出 `fallback`，无 fallback 时原样输出。

### date / datetime

```ts
{ type: "date" }                       // 默认 yyyy-MM-dd
{ type: "datetime", pattern: "yyyy-MM-dd HH:mm:ss" }
```

接受 `Date` / 可解析字符串 / 时间戳。Workbook 路径写入 Excel 日期序列并自动注入 `numFormat`；Stream 路径输出 pattern 格式化字符串。统一按 **UTC 分量**解释（与 Workbook 序列的 `dateToSerial` 口径一致，跨路径/跨时区显示相同）；ISO 日期字符串按 ECMA-262 解析为 UTC 午夜，详见[值格式化的时区约定](/zh/packages/excel-exporter/guide/04-formatting)。

### number

```ts
{ type: "number", decimals: 2, thousands: true }
```

`decimals` 默认 0，`thousands` 默认 false。**务必显式声明 `decimals`**：Workbook 路径保留完整精度经 `numFormat` 渲染，Stream/兜底路径将 `decimals` 烧入存储值，两种路径存储值可能不同。

`thousands` 的跨路径差异：Workbook 路径经自动注入的 `#,##0` `numFormat` 渲染千分位；Stream 路径（≥ 50,000 行 / 降级导出）无法使用 `numFormat`，单元格保持为**数字**，因此不显示千分位（把分隔符烧入值会把数据单元格变成文本，破坏下游计算）。

`null`/`undefined` 在所有路径下均渲染为空单元格——绝不会是 `0`。

### padding

```ts
{ type: "padding", fill: "0", length: 6 } // "42" -> "000042"
```

省略 `align`（默认）对应 `padStart`，在**左侧**补字符（值右对齐），适合工号、订单号的前导零场景；`align: "left"` 对应 `padEnd`，在**右侧**补字符（值左对齐）。

## 函数形式

```ts
format: (value, row) => string | number | boolean;
```

可以访问整行数据做条件格式化。函数无法穿过结构化克隆，因此各路径行为不同：

- **main 路径**（auto 模式下浏览器 < 20,000 行 / Node < 50,000 行；或任意行数显式 `mode: "main"`）：函数正常执行；
- **Node 的 stream 路径**（≥ 50,000 行）：同样在主线程执行，函数正常执行；
- **浏览器 worker 路径**（auto ≥ 20,000 行，或显式 `mode: "worker"` / `mode: "stream"`）：函数会被**剥离并打印 `console.warn`**，该列以原始值导出（不会报错，也不会回落到 main）。

需要 worker 路径保留格式时，请改写为 FormatSpec。
