---
"@marcusok/excel-exporter": patch
---

Fix the progress overlay staying invisible when a new export starts during the previous overlay's fade-out: the reused node kept `opacity: 0` (from the cancelled teardown) for the whole run while still blocking page interaction. The overlay now fades back in and resets to its initial indeterminate state on reuse.

Also:

- `configureWasm({ workerUrl })` now prints a warning when the worker URL changes: the shared Worker reads its script URL once at creation, so a later change only reaches a worker created after `terminateWorker()` (the `wasmUrl` caveat already warned; `workerUrl` did not).
- The reserved `__index__` prop check is now alias-aware on every path (conflict detection, Workbook and Fast stream builders): a hand-written legacy `key: "__index__"` column is rejected when `indexColumn` is enabled and driven by the row number on the low-level builders, instead of being silently read from `data`.
