---
"@marcusok/excel-exporter": patch
---

Docs: note that `parseWorkerUrl` (present on the shared `LoaderOptions` type behind `configureWasm`) configures `@marcusok/excel-preview`'s parse worker and has no effect on this package — the export worker's URL field is `workerUrl`.
