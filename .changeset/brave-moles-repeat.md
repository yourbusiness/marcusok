---
"@marcusok/excel-preview": patch
---

Fix formatting and alignment fidelity issues:

- Cells with an explicit `horizontal="general"` alignment now fall back to type-based alignment (numbers render right-aligned), matching Excel semantics for the default alignment.
- A workbook whose `activeTab` points at a hidden sheet now opens the first visible sheet instead of silently rendering hidden content with no tab to switch away.
- Literal-only number formats (e.g. `"yes"`) no longer append the numeric value — Excel hides the number when a section has no digit placeholder.
- Number formats whose sections are all conditional with no match (e.g. `[>100]0.00` for 5) now render empty instead of falling back to the first conditional section.
- Font names are sanitized more strictly (control characters stripped) when compiling the built-in stylesheet, closing a CSS-injection vector from untrusted xlsx files.
