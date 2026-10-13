# API: FormatSpec

A structured, thread-safe value-formatting description. Prefer FormatSpec on worker/stream paths; the actual reach of the function form is covered in "Function form" at the end.

## Type definition

```ts
type FormatSpec =
  | { type: "enum"; map: Record<string, string>; fallback?: string }
  | { type: "date"; pattern?: string } // default "yyyy-MM-dd"
  | { type: "datetime"; pattern?: string } // default "yyyy-MM-dd HH:mm"
  | { type: "number"; decimals?: number; thousands?: boolean }
  | { type: "padding"; fill: string; length: number; align?: "left" | "right" };
```

## Per type

### enum

```ts
{ type: "enum", map: { paid: "Paid", pending: "Pending" }, fallback: "Unknown" }
```

Outputs the mapped label; unmapped values use `fallback`, or pass through when absent. Missing values — `null`/`undefined` and blank/whitespace-only strings — render as **empty cells**: they never consult `map` nor fall to `fallback` (a configured `fallback` covers present-but-unmapped values, not data holes, so an export never fills them with business-meaningful fake text).

### date / datetime

```ts
{ type: "date" }                       // default yyyy-MM-dd
{ type: "datetime", pattern: "yyyy-MM-dd HH:mm:ss" }
```

Accepts `Date` / parseable string / timestamp. The Workbook path writes an Excel date serial and auto-injects `numFormat`; the stream path outputs the pattern-formatted string. Values are interpreted by their **UTC components** (matching the workbook serial's `dateToSerial` convention, so all paths and timezones render the same); ISO date strings parse as UTC midnight per ECMA-262 — see the timezone note in [Value Formatting](/packages/excel-exporter/guide/04-formatting). Pattern tokens differ between paths: the stream path lower-cases the pattern, parses only `yyyy` / `MM` / `dd` / `HH` / `mm` / `ss`, and emits anything else as-is — not interpreting quoted literals, so `yyyy"年"M"月"` renders `2026"年"m"月"` — while the Workbook path renders any valid Excel format code, quoted literals included — see the note in [Value Formatting](/packages/excel-exporter/guide/04-formatting).

Two more edges in the same area: a value the parser cannot convert to a date (bigints, booleans, plain objects, invalid Dates — only `Date` / parseable strings / timestamps are accepted) is written as its string form instead of failing the export; and superset tokens (`mmm`, `yy`, single-letter `M`/`d`) are only partially handled on the stream path — `"mmm"` renders as `"09m"` there while Excel renders the month abbreviation — so stick to the six tokens for cross-threshold consistency (details in [Value Formatting](/packages/excel-exporter/guide/04-formatting)).

### number

```ts
{ type: "number", decimals: 2, thousands: true }
```

`decimals` defaults to 0, `thousands` to false; `decimals` is validated as an integer in 0–100 (the `toFixed` ceiling the stream path relies on). **Always set `decimals` explicitly**: the Workbook path keeps full precision rendered via `numFormat`, while Stream/fallback paths bake decimals into the stored value — the two can differ otherwise.

Cross-path `thousands`: the Workbook path renders the separator via an auto-injected `#,##0` `numFormat`; the stream path (>= 50,000 rows / degraded exports) cannot use `numFormat` and keeps the cell a **number**, so separators are not visible there (baking them into the value would turn data cells into text and break downstream calculations).

`null`/`undefined` values (and blank/whitespace-only strings — the most common missing-value shape coming out of databases, forms and CSV imports) render as empty cells on every path — never `0`. `padding` treats them the same way: an empty cell, not a padded fake code like `"00000"`. `enum` does too: an empty cell, never the `fallback` (see above).

### padding

```ts
{ type: "padding", fill: "0", length: 6 } // "42" -> "000042"
```

Omitting `align` (the default) maps to `padStart`: the fill goes on the **left** (value right-aligned) — right for leading-zero IDs. `align: "left"` maps to `padEnd`: the fill goes on the **right** (value left-aligned). `length` is validated as an integer in 0–10000 (an out-of-range value fails up front with a clear error instead of blowing up — or silently not padding — at render time).

## Up-front validation

The whole spec family is validated before any route runs — a violation fails the export with `{ success: false }` and an error naming the column, instead of surfacing later as an opaque error or a silently degraded file. Beyond the `decimals` / `length` ranges above:

- `type` must be exactly one of the five literals — a typo like `"percent"` is rejected up front (otherwise the value would silently stringify, turning a numeric column into text with no warning);
- `enum.map` must be a non-array, non-null object (missing / `null` / an array are all rejected);
- a provided `date` / `datetime` `pattern` must be a string;
- `padding.fill` must be a string — it is required on the type, so `undefined` is rejected as well (a non-string would not throw later but produce odd output, hence the up-front gate).

## Function form

```ts
format: (value, row) => string | number | boolean;
```

Can access the whole row for conditional formatting. Functions cannot cross the structured-clone boundary, so behavior differs per path:

- **main path** (auto mode: browser < 20,000 rows / Node < 50,000 rows; or any size with explicit `mode: "main"`): executed normally;
- **Node's stream path** (≥ 50,000 rows): also main-thread, executed normally;
- **main-thread retry after a worker failure**: executed normally — the original options keep the function, only the copy sent to the worker is stripped;
- **browser worker path** (auto ≥ 20,000 rows, or explicit `mode: "worker"` / `mode: "stream"`): functions are **stripped with a `console.warn`** and the column exports its raw value (no error, no fallback to main).

Convert to FormatSpec to keep formatting on the worker path.
