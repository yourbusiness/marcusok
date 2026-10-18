---
"@marcusok/excel-exporter": patch
---

Validate Excel's hard sheet limits (1,048,576 rows including the header block, 16,384 columns) up front on every path. Over-limit input previously zipped successfully on the stream path and produced a workbook Excel flags as corrupt while reporting `success: true`; it now fails fast with a structured `{ success: false, error }` and no fallback attempt.
