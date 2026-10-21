# 值格式化

列级 `format` 可以是**结构化 FormatSpec**（跨线程安全，所有路径可用）或**函数**（主线程路径执行；浏览器 worker 路径会被剥离，见下文「函数形式」）。

## FormatSpec

| 类型       | 参数                                     | 示例                                                          | 说明                                                                                |
| ---------- | ---------------------------------------- | ------------------------------------------------------------- | ----------------------------------------------------------------------------------- |
| `enum`     | `map`、`fallback?`                       | `{ type: "enum", map: { paid: "已支付" }, fallback: "未知" }` | 枚举值映射；未命中且无 fallback 时原样输出；缺失值（null/空白）渲染为空单元格       |
| `date`     | `pattern?`（默认 `yyyy-MM-dd`）          | `{ type: "date" }`                                            | 转 Excel 日期序列并自动注入 `numFormat`                                             |
| `datetime` | `pattern?`（默认 `yyyy-MM-dd HH:mm`）    | `{ type: "datetime" }`                                        | 同上，带时间                                                                        |
| `number`   | `decimals?`（默认 0）、`thousands?`      | `{ type: "number", decimals: 2, thousands: true }`            | 数字语义：两条路径都保留完整精度并经 numFmt 渲染；缺失值（null/空白）渲染为空单元格 |
| `padding`  | `fill`、`length`、`align?`（left/right） | `{ type: "padding", fill: "0", length: 6, align: "left" }`    | 左/右补全到固定长度（如工号）                                                       |

```ts
columns: [
  { prop: "orderId", label: "订单号", width: 12 },
  {
    prop: "date",
    label: "日期",
    width: 12,
    format: { type: "date", pattern: "yyyy/MM/dd" },
  },
  {
    prop: "amount",
    label: "金额",
    width: 14,
    format: { type: "number", decimals: 2, thousands: true },
  },
  {
    prop: "status",
    label: "状态",
    width: 10,
    format: { type: "enum", map: { paid: "已支付" }, fallback: "未知" },
  },
  {
    prop: "code",
    label: "编码",
    width: 12,
    format: { type: "padding", fill: "0", length: 6, align: "right" },
  },
];
```

## 函数形式

```ts
{
  prop: "amount",
  label: "金额",
  width: 14,
  format: (value, row) => {
    const n = Number(value);
    return n >= 1000 ? `大额 ${n.toFixed(2)}` : n.toFixed(2);
  },
}
```

函数签名：`(value: unknown, row: Record<string, unknown>) => string | number | boolean`。函数无法穿过结构化克隆，各路径行为不同：main 路径（浏览器低于 Worker 门槛 / Node 低于 stream 门槛，或任意行数下显式 `mode: "main"`）、Node 的 stream 路径（auto 达到 stream 门槛或显式 `mode: "stream"`，同样在主线程执行）、以及 worker 失败后的主线程重试（原始 options 里的函数仍在——被剥离的只是发给 worker 的副本）都正常执行；浏览器 worker 路径（auto 达到 Worker 门槛，或显式 `mode: "worker"` / `mode: "stream"`）会**剥离函数并打印 `console.warn`**，该列以原始值导出（不报错、不回落 main）。需要 worker 路径保留格式时改为 FormatSpec。

函数抛错不会被隔离：异常会从构建过程抛出，并可能触发降级/失败链，因此请在函数内部处理坏值。返回 `null` / `undefined`（或空串）会渲染为空单元格。

## 跨模式精度注意事项

两条路径现在共享同一套存储语义（Stream 路径会为 `FormatSpec` 输出最小 `styles.xml`）：

- **存储值**：`number` 在两条路径都保留完整精度；`decimals`/`thousands` 是纯显示 numFmt。`date`/`datetime` 在两条路径存同一个 Excel 序列号。
- **显示**：Workbook 路径通过单元格样式附加 numFmt；Stream 路径为每个不同格式码输出一个 cellXf。跨 50,000 行不再改变单元格类型或存储值。
- **单元格/布局样式**（字体、填充、边框、宽度、冻结、筛选）仍仅 Workbook 路径支持；Stream 只恢复由 FormatSpec 推导的数字/日期格式。
- **缺失值在所有路径渲染为空单元格**，`number`、`padding` 与 `enum` 一视同仁：`null`/`undefined` 与空串/纯空白串（数据库、表单、CSV 导入中最常见的缺失值形态）绝不会变成 `0`、`"00000"` 或 `enum` 的 `fallback`。

## 日期值

`date` / `datetime` 接受 `Date` 对象、可解析字符串或时间戳。两条路径都写入日期序列号，并把 pattern 作为 numFmt 附加（Workbook：单元格样式；Stream：最小 `styles.xml`），因此单元格类型与显示跨阈值一致——**前提是没有同时设置冲突的 `CellStyle.numFormat`**。Workbook 路径上显式的 `style.numFormat` 优先于 FormatSpec 推导出的格式码；Stream 路径忽略 `CellStyle.numFormat`，始终使用 FormatSpec pattern。保持两者一致（或省略 `style.numFormat`）才能维持跨路径显示一致。

**pattern**：两条路径都把 pattern 作为数字格式交给 Excel 渲染，`yy`、单字母 `m`/`d`、`AM/PM`、引号字面量等完整格式码语义一致，不再有六 token 子集限制。

**时区约定（跨路径一致）**：`date` / `datetime` 统一按值的 **UTC 分量**解释与输出——两条路径的序列都来自 `dateToSerial`（UTC 口径），因此同一输入在任何时区、任何路径下显示一致；ISO 日期字符串（如 `"2026-07-01"`）按 ECMA-262 解析为 UTC 午夜，天然符合该口径。注意 `new Date(年, 月, 日)` 这类**本地时间**构造的 Date，其 UTC 分量在非零时区可能落到前一天（例如 UTC+8 的本地 0 点 = 前一日 16:00 UTC）。为保证跨时区一致，日期列建议传 ISO 字符串或用 `Date.UTC(...)` 构造。
