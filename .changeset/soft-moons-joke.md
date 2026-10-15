---
"@marcusok/excel-preview": patch
---

Tightened the text-format sniff: a mostly-non-ASCII binary whose printable ratio passes the 95% heuristic (e.g. a 0xFF-filled blob) is no longer rendered as garbled CSV — the sampled bytes must now decode as UTF-8 or GB18030, otherwise the parse fails with `UNSUPPORTED`. GBK/GB18030 CSVs, including ones truncated mid-character at the 2KB sniff window, still parse as before. Also synced the stale `PreviewOptions.sheet` JSDoc with the actual behavior (an invalid initial sheet reports `Sheet not found` through `onError` before falling back to `activeTab`).
