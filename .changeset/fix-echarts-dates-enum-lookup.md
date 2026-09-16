---
"@marcusok/excel-exporter": patch
---

Robustness fixes from a code audit. Invalid Date values no longer crash every export path with a raw `RangeError` (they render as `"Invalid Date"` strings, matching `String(new Date(NaN))`). Enum FormatSpec lookups no longer walk the prototype chain: values like `"constructor"` / `"__proto__"` that the map does not define now hit the fallback instead of writing `function Object() {...}` text into the cell. Object-form ECharts scatter datums (`{ value: [x, y] }`) are now exported as X/Y pairs exactly like bare `[x, y]` pairs instead of being silently stringified as `"[1,2]"`. A missing or empty `filename` fails fast with a structured `{ success: false }` error instead of returning `success: true` with no file and a cryptic warning. The browser download trigger releases its object URL even when a sandboxed DOM throws mid-trigger.
