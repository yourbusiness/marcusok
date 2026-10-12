---
"@marcusok/excel-preview": patch
---

Number-format fidelity fixes for zh-CN Excel files and Excel-exact rounding:

- Builtin `numFmtId` 27–36 and 50–58 (East-Asian groups, only written by East-Asian-locale Excel/WPS) now resolve to zh-CN behavior strings — zh-CN dates stored as builtin ids (e.g. id 31 for `2025年1月21日`) render as dates instead of raw serials like `45678`.
- `[DBNum1/2/3]` digit-substitution modifiers are stripped so the numeric section still renders; unknown bracket keywords no longer misread the section as a date (`[DBNum2][$-804]#,##0.00` used to render as `[18BNu52]#,##0.00`).
- Display rounding now follows Excel's decimal semantics (15 significant digits, half-up): `2.675` with `0.00` shows `2.68` (the engine's binary `toFixed` produced `2.67`); percentages round in the display domain (`0.12345` with `0.00%` → `12.35%`).
- Trailing-comma thousands scaling is supported (`#,##0,` divides by 1000 per trailing comma, `0.0,,"M"` → `1.5M`).
- Section-level `General` (`[Red]General`, negative-section `General;-General`, WPS's `G/通用格式`) renders the value instead of the literal text "General".
