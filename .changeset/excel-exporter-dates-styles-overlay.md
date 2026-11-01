---
"@marcusok/excel-exporter": patch
---

Handle cross-realm `Date` values (iframe / `node:vm`) as dates on both the Workbook and Fast stream paths instead of JSON-quoting them; generate styled blank cells for missing values so `dataStyle` borders/fills no longer break on sparse rows; and apply the latest caller's overlay theme/zIndex/blockInteraction plus a fresh `minVisibleMs` window when the shared overlay node is reused. The overlay no longer waits two animation frames unless `delayMs: 0`, removing a fixed 16–33ms cost from every default export.
