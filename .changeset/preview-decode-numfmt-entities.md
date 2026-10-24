---
"@marcusok/excel-preview": patch
---

Fix custom number formats with quoted literals (e.g. `yyyy"年"m"月"d"日"`, `0"元"`) rendering as `&quot;` garbled text. The engine returns the raw `styles.xml` attribute value without XML entity decoding; the model layer now decodes `formatCode` before resolving it into `PreviewXf.numFmtCode`.
