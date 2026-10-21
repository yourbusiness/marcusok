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

Accepts `Date` / parseable string / timestamp. Both paths write an Excel date serial and attach the pattern as a numFmt — the Workbook path through the cell style, the stream path through its minimal `styles.xml` — so the stored value and the displayed date agree across routes (subject to the explicit `style.numFormat` caveat below). Values are interpreted by their **UTC components** (matching the workbook serial's `dateToSerial` convention, so all paths and timezones render the same); ISO date strings parse as UTC midnight per ECMA-262 — see the timezone note in [Value Formatting](/packages/excel-exporter/guide/04-formatting). The pattern is handed to Excel as a number format on both paths, so the full format-code language applies (quoted literals included); there is no reduced token subset. The standalone `formatDateByPattern` helper (exported for custom consumers and used by `displayValue`, but not by the stream export) is the one with the reduced, lowercase-only token scan — see its API entry.

One more edge in the same area: a value the parser cannot convert to a date (bigints, booleans, plain objects, invalid Dates — only `Date` / parseable strings / timestamps are accepted) is written as its string form instead of failing the export.

### number

```ts
{ type: "number", decimals: 2, thousands: true }
```

`decimals` defaults to 0, `thousands` to false; `decimals` is validated as an integer in 0–100. Both paths keep the full-precision stored value and render `decimals`/`thousands` through an Excel numFmt — the Workbook path via the cell style, the stream path via its minimal `styles.xml`. Crossing 50,000 rows therefore no longer rounds the stored value or drops separators.

Dates/datetimes behave the same way: both paths store the Excel serial number (UTC components) and attach the pattern as a numFmt, so the cell type and display agree across routes.

**Explicit `CellStyle.numFormat` vs FormatSpec.** On the Workbook path a column's explicit `style.numFormat` takes precedence and the FormatSpec-derived code is not injected. The stream path drops `CellStyle.numFormat` entirely and uses the FormatSpec pattern/decimals/thousands. If both are set and differ, the stored value still matches but the display differs across the 50k threshold; keep them consistent or omit `style.numFormat` on FormatSpec columns.

`null`/`undefined` values (and blank/whitespace-only strings — the most common missing-value shape coming out of databases, forms and CSV imports) render as empty cells on every path — never `0`. `padding` treats them the same way: an empty cell, not a padded fake code like `"00000"`. `enum` does too: an empty cell, never the `fallback` (see above).

### padding

```ts
{ type: "padding", fill: "0", length: 6 } // "42" -> "000042"
```

Omitting `align` (the default) maps to `padStart`: the fill goes on the **left** (value right-aligned) — right for leading-zero IDs. `align: "left"` maps to `padEnd`: the fill goes on the **right** (value left-aligned). Padding never truncates: a value already at least `length` characters long is returned unchanged. `length` is validated as an integer in 0–10000 (an out-of-range value fails up front with a clear error instead of blowing up — or silently not padding — at render time).

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
- **Node's stream path** (auto ≥ 50,000 rows, or explicit `mode: "stream"`): also main-thread, executed normally;
- **main-thread retry after a worker failure**: executed normally — the original options keep the function, only the copy sent to the worker is stripped;
- **browser worker path** (auto ≥ 20,000 rows, or explicit `mode: "worker"` / `mode: "stream"`): functions are **stripped with a `console.warn`** and the column exports its raw value (no error, no fallback to main).

Convert to FormatSpec to keep formatting on the worker path.

Function errors are **not** isolated: a throw propagates out of the build and can trigger the normal degradation/failure path, so make the function total and handle bad values yourself. Returning `null` / `undefined` (or an empty string) renders an empty cell.
