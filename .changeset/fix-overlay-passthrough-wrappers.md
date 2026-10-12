---
"@marcusok/excel-exporter": patch
---

Forward the `overlay` option through `exportTable` and `exportEcharts`: both options types now accept `overlay?: boolean | ProgressOverlayOptions` and the adapters pass it to `exportExcel`, so the default progress overlay (and `overlay: false`) behaves identically across all three entry points as documented.
