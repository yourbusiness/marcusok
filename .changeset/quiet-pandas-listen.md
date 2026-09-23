---
"@marcusok/excel-exporter": patch
---

Treat blank/whitespace-only strings as missing values in `{ type: "padding" }` columns, aligning the behavior with the documented missing-value semantics: they now render as empty cells on every path instead of being padded into a fake-looking code like `"00000"` (previously only `null`/`undefined` were guarded). The `FormatSpec` JSDoc's note about which lookup key `{ type: "enum" }` uses for missing values is also more precise now.
