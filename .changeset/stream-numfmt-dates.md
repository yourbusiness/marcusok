---
"@marcusok/excel-exporter": minor
---

Keep date and number semantics identical across the 50k-row threshold. The fast stream now emits a minimal `styles.xml` for `FormatSpec` columns: dates/datetimes are stored as Excel serial numbers plus a numFmt (instead of formatted text), and `decimals`/`thousands` are display-only numFmt settings rather than being baked into stored values. Cell/layout styling remains Workbook-only.
