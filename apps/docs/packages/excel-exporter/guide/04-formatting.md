# Value Formatting

Column `format` can be a structured **FormatSpec** (thread-safe; works on all paths) or a **function** (executed on main-thread paths; stripped on the browser worker path — see "Function form" below).

## FormatSpec

| Type       | Options                                 | Example                                                        | Description                                                                                                                  |
| ---------- | --------------------------------------- | -------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `enum`     | `map`, `fallback?`                      | `{ type: "enum", map: { paid: "Paid" }, fallback: "Unknown" }` | Maps raw values to labels; unmapped values use `fallback` or pass through; missing values (null/blank) render as empty cells |
| `date`     | `pattern?` (default `yyyy-MM-dd`)       | `{ type: "date" }`                                             | Converts to an Excel date serial and auto-injects `numFormat`                                                                |
| `datetime` | `pattern?` (default `yyyy-MM-dd HH:mm`) | `{ type: "datetime" }`                                         | Same, with time                                                                                                              |
| `number`   | `decimals?` (default 0), `thousands?`   | `{ type: "number", decimals: 2, thousands: true }`             | Numeric semantics; both paths keep full precision and render via numFmt; missing values (null/blank) render as empty cells   |
| `padding`  | `fill`, `length`, `align?` (left/right) | `{ type: "padding", fill: "0", length: 6, align: "left" }`     | Pads to a fixed length (IDs, codes)                                                                                          |

```ts
columns: [
  { prop: "orderId", label: "Order ID", width: 12 },
  {
    prop: "date",
    label: "Date",
    width: 12,
    format: { type: "date", pattern: "yyyy/MM/dd" },
  },
  {
    prop: "amount",
    label: "Amount",
    width: 14,
    format: { type: "number", decimals: 2, thousands: true },
  },
  {
    prop: "status",
    label: "Status",
    width: 10,
    format: { type: "enum", map: { paid: "Paid" }, fallback: "Unknown" },
  },
  {
    prop: "code",
    label: "Code",
    width: 12,
    format: { type: "padding", fill: "0", length: 6, align: "right" },
  },
];
```

## Function form

```ts
{
  prop: "amount",
  label: "Amount",
  width: 14,
  format: (value, row) => {
    const n = Number(value);
    return n >= 1000 ? `Large ${n.toFixed(2)}` : n.toFixed(2);
  },
}
```

Signature: `(value: unknown, row: Record<string, unknown>) => string | number | boolean`. Functions cannot cross the structured-clone boundary, so behavior differs per path: the main path (browser < 20,000 rows / Node < 50,000 rows, or any size with explicit `mode: "main"`), Node's stream path (auto ≥ 50,000 rows or explicit `mode: "stream"`, also main-thread) and the main-thread retry after a worker failure (the original options keep the function — only the copy sent to the worker is stripped) execute them normally; the browser worker path (auto ≥ 20,000 rows, or explicit `mode: "worker"` / `mode: "stream"`) **strips them with a `console.warn`** and exports the raw value (no error, no fallback to main). Convert to FormatSpec to keep formatting on the worker path.

A thrown function is not isolated: it propagates out of the build and may trigger the degradation/failure path, so handle bad values inside the function. Returning `null` / `undefined` (or an empty string) renders an empty cell.

## Cross-path precision notes

Both routes now share the same storage semantics — the stream path emits a minimal `styles.xml` for `FormatSpec` formats:

- **Stored values**: `number` keeps full precision on both paths; `decimals`/`thousands` are display-only numFmt settings. `date`/`datetime` store the same Excel serial number on both paths.
- **Display**: the Workbook path attaches the numFmt via the cell style; the stream path emits one cellXf per distinct format code. Crossing 50,000 rows no longer changes cell type or stored value.
- **Cell/layout styling** (font, fill, border, width, freeze, filter) is still Workbook-only; only FormatSpec-derived number/date formats are restored on the stream.
- **Missing values render as empty cells** on every path, for `number` and `padding` columns and `enum` alike: `null`/`undefined` and blank/whitespace-only strings (the most common missing-value shape from databases, forms and CSV imports) never become `0`, a padded `"00000"`, or the `enum` `fallback`.

## Dates

`date` / `datetime` accept `Date` objects, parseable strings or timestamps. Both paths store an Excel serial and attach the pattern as a numFmt (Workbook: cell style; Stream: minimal `styles.xml`), so cell type and display agree across the 50k threshold — **provided you do not also set a conflicting `CellStyle.numFormat`**. On the Workbook path an explicit `style.numFormat` wins over the FormatSpec-derived code; the stream path ignores `CellStyle.numFormat` and always uses the FormatSpec pattern. Keeping the two equal (or omitting `style.numFormat`) preserves cross-path display parity.

**Pattern semantics**: both paths hand the pattern to Excel as a number format (`numFormat` / minimal `styles.xml`), so the full format-code language applies identically — `yy`, single-letter `m`/`d`, `AM/PM`, `mmm`/`dddd`, quoted literals like `yyyy"年"M"月"d"日"`. There is no reduced token subset on the stream path anymore.

**Timezone convention (consistent across paths)**: `date` / `datetime` interpret and render values by their **UTC components** — both paths store the serial produced by modern-xlsx's `dateToSerial` (UTC wall clock), so one input renders identically on every path in every timezone. ISO date strings (`"2026-07-01"`) parse as UTC midnight per ECMA-262 and fit this convention natively. Note that Dates constructed with local time (`new Date(year, month, day)`) carry UTC components that can fall on the previous day in non-UTC timezones (local midnight in UTC+8 = 16:00 UTC the day before). Prefer ISO strings or `Date.UTC(...)` for timezone-stable date columns.
