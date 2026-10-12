---
"@marcusok/excel-exporter": patch
---

Fix `enum` FormatSpec rendering the `fallback` text for missing values. `null`, `undefined`, empty and whitespace-only strings now render as empty cells — matching the `number` and `padding` branches — instead of being filled with the fallback label (e.g. "Unknown"), which turned data holes into meaningful-looking fake values.
