---
"@marcusok/excel-preview": patch
---

An invalid initial `sheet` option (unknown name, out-of-range or non-integer index) now reports `Sheet not found: …` through `onError` — same wording and code as the instance method `setSheet()` — instead of silently falling back. The fallback rendering behavior is unchanged: the preview still renders the file's `activeTab`.
