---
"@marcusok/excel-exporter": minor
---

- fix: the Workbook path (default route under 50k rows) now writes `NaN`/`Infinity` as visible strings instead of illegal `<v>NaN</v>` XML that Excel flags as corrupt, matching the stream and SheetJS paths.
- fix: `exportExcel` now fails invalid input (bad merges, duplicate/invalid sheet names, missing column keys) immediately via a pre-flight check, instead of first attempting a pointless SheetJS fallback — same error messages as before, no misleading "Falling back" warning.
- The SheetJS fallback result's `error` message now carries the degradation reason programmatically (e.g. `workerUrl not configured`).
- feat: export the `LoaderOptions` / `LoadState` types from the package entry (documented but previously missing).
- docs: document the SheetJS npm CVE situation; installing the optional `xlsx` peer from the official CDN tarball is now recommended.
