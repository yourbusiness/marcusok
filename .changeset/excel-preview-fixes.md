---
"@marcusok/excel-preview": patch
---

Parse namespace-prefixed and single-quoted `styles.xml` (with a count-based fallback to engine styles), stop treating "PK"-prefixed CSV as ZIP, choose the CSV delimiter by per-row consistency, display serial 0 as Excel's `1900-01-00`, accept cross-realm byte sources and async callbacks, and correct the timeout/gradient documentation.
