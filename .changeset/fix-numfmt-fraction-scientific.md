---
"@marcusok/excel-preview": patch
---

Fix number-format rendering for fractions and scientific notation: the display-rounding pre-pass no longer pre-rounds fraction cores (`# ?/?`, `# ??/??` — previously `1.25` rendered as `1`) or scientific cores (`0.00E+00` — previously tiny values like `0.00012` collapsed to `0.00E+0`), and scientific exponents are normalized to Excel's two-digit signed form (`1.23E+4` → `1.23E+04`).
