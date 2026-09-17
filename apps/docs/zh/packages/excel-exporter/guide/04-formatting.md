# 值格式化

列级 `format` 可以是**结构化 FormatSpec**（跨线程安全，所有路径可用）或**函数**（主线程路径执行；浏览器 worker 路径会被剥离，见下文「函数形式」）。

## FormatSpec

| 类型       | 参数                                     | 示例                                                          | 说明                                                     |
| ---------- | ---------------------------------------- | ------------------------------------------------------------- | -------------------------------------------------------- |
| `enum`     | `map`、`fallback?`                       | `{ type: "enum", map: { paid: "已支付" }, fallback: "未知" }` | 枚举值映射；未命中且无 fallback 时原样输出               |
| `date`     | `pattern?`（默认 `yyyy-MM-dd`）          | `{ type: "date" }`                                            | 转 Excel 日期序列并自动注入 `numFormat`                  |
| `datetime` | `pattern?`（默认 `yyyy-MM-dd HH:mm`）    | `{ type: "datetime" }`                                        | 同上，带时间                                             |
| `number`   | `decimals?`（默认 0）、`thousands?`      | `{ type: "number", decimals: 2, thousands: true }`            | 数字语义：Workbook 路径保留完整精度，经 `numFormat` 渲染 |
| `padding`  | `fill`、`length`、`align?`（left/right） | `{ type: "padding", fill: "0", length: 6, align: "left" }`    | 左/右补全到固定长度（如工号）                            |

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

函数签名：`(value: unknown, row: Record<string, unknown>) => string | number | boolean`。函数无法穿过结构化克隆，各路径行为不同：main 路径（浏览器 < 20,000 行 / Node < 50,000 行）、Node 的 stream 路径（≥ 50,000 行，同样在主线程执行）、以及 worker 失败后的主线程重试（原始 options 里的函数仍在——被剥离的只是发给 worker 的副本）都正常执行；浏览器 worker 路径（auto ≥ 20,000 行，或显式 `mode: "worker"` / `mode: "stream"`）会**剥离函数并打印 `console.warn`**，该列以原始值导出（不报错、不回落 main）。需要 worker 路径保留格式时改为 FormatSpec。

## 跨模式精度注意事项

同一份配置在不同路径下行为有细微差异，务必显式声明 `decimals`：

- **Workbook 路径**（main / worker+workbook）：`number` 保留完整精度，显示小数由自动注入的 `numFormat` 控制；
- **Stream 路径**（≥ 50,000 行）：没有 `numFormat` 支持，会把 `decimals` 烧进存储值（如 `9999.99` → `10000`）；
- 因此**不声明 `decimals`（默认 0）时，两种路径存进单元格的值可能不同**；跨阈值一致性是显式声明 `decimals` 的最重要理由。

## 日期值

`date` / `datetime` 接受 `Date` 对象、可解析字符串或时间戳。Workbook 路径写入日期序列 + `numFormat`；Stream 路径（无 `numFormat` 支持）会按 pattern 输出可读字符串（`mm` 会自动按前后文区分月份与分钟）。

**pattern 支持的 token（跨路径有差异）**：Stream 路径（≥ 50,000 行、显式 `mode: "stream"` 或兜底路径）只解析 `yyyy` / `MM` / `dd` / `HH` / `mm` / `ss` 六种 token（大小写不敏感）。Workbook 路径把 pattern 作为 `numFormat` 交给 Excel 渲染，任意合法格式码都生效（`yy`、单字母 `m`/`d`、`AM/PM`、`yyyy"年"` 这类字面量等）。六种 token 之外的**字符**在 Stream 路径按字面原样输出——比如 `pattern: "yy-MM-dd"`，阈值以下导出正常的两位年份，阈值以上导出 `yy-01-05`。注意超集 token 并非整体原样：其中的六 token 前缀仍会被解析、剩余字符漏出（`"mmm"` → `"09m"`，而 Excel 侧渲染月份缩写）。为保证跨阈值一致，请只使用这六种 token。

**时区约定（跨路径一致）**：`date` / `datetime` 统一按值的 **UTC 分量**解释与输出——Workbook 路径的序列来自 modern-xlsx 的 `dateToSerial`（UTC 口径），Stream 路径的字符串同样取 UTC 分量，因此同一输入在任何时区、任何路径下显示一致；ISO 日期字符串（如 `"2026-07-01"`）按 ECMA-262 解析为 UTC 午夜，天然符合该口径。注意 `new Date(年, 月, 日)` 这类**本地时间**构造的 Date，其 UTC 分量在非零时区可能落到前一天（例如 UTC+8 的本地 0 点 = 前一日 16:00 UTC）。为保证跨时区一致，日期列建议传 ISO 字符串或用 `Date.UTC(...)` 构造。
