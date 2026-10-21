---
"@marcusok/excel-exporter": minor
"@marcusok/excel-preview": minor
---

Harden export input handling, routing and preview dimensions.

- `excel-exporter`: reject XML 1.0-illegal characters in sheet names, font names, number formats and date/datetime patterns (they previously produced workbooks Excel flags as corrupt while the export still returned `success: true`); sanitize data/header text consistently on both routes; runtime-validate style enums and types for JS callers; include the estimated cell count in auto routing so wide sheets no longer take the slow Workbook path; expose dropped fast-stream features as `ExportResult.degradedFeatures`; inject overlay styles/nodes into the container's own document (iframe support).
- `excel-preview`: reject out-of-range row/column/merge dimensions as `CORRUPT` before the main-thread layout allocates geometry arrays; normalize negative/non-finite row heights and column widths; inject renderer styles and create DOM nodes in the container's own document (iframe support).
