---
"@marcusok/excel-preview": patch
---

Fix a batch of number-format fidelity bugs, each verified against Excel semantics with regression tests:

- `[Color n]` now resolves through Excel's legacy 56-color palette (1-based) instead of treating `n` as a 0-based indexed-palette index — the whole group was shifted by one (`[Color 1]` rendered white while `[Black]` rendered black, `[Color 3]` green instead of red); values out of 1..56 produce no color.
- Fraction formats respect the denominator placeholder width: `# ?/?` picks the closest fraction with a denominator ≤ 9 (`3.14159` → `3 1/7`, previously `3 14/99`), reduces to lowest terms, and a fraction format without an integer placeholder renders an improper fraction (`?/?` + `1.25` → `5/4`).
- `mm:ss.0`-style fractional seconds round the whole time base at display precision and carry into the minute — `59.99999s` renders `1:00.0` instead of the illegal `01:60.0`, and `59.94s` stays `0:59.9` instead of wrongly carrying to `1:00`.
- Conditional sections accept scientific-notation thresholds and leading signs (`[>1E+10]`), which previously leaked the bracket into the output as literal garbage.
- Section-level `General` with edge literals keeps the value: `General" kg"` renders `5 kg` instead of dropping the number and leaking the word "General".
- A text section with multiple `@` tokens substitutes every `@` (`@@@` + `abc` → `abcabcabc`); previously only the first was replaced and the rest were dropped.
- Consecutive percent signs each multiply by 100 (`0.00%%` + `0.001` → `10.00%%`), matching Excel.
- `am/pm` and `a/p` mirror the case of the format code (`h:mm am/pm` + noon → `12:00 pm`, previously always `PM`).
