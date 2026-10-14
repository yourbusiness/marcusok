---
"@marcusok/excel-exporter": minor
---

A worker timeout now fails fast instead of retrying on the main thread. When the shared export worker exceeds `workerTimeoutMs` (default 120s), it is terminated and the export resolves with `success: false` carrying the original timeout error — no main-thread retry, no style-less stream fallback. Requests queued behind the timed-out export are rejected with the same fast-fail semantics. Rationale (aligned with `@marcusok/excel-preview`'s parse timeout): the timeout means the worker is wedged or the workload legitimately exceeds the budget, and rerunning the same large build on the main thread is exactly the freeze the worker route exists to avoid. Retry the export (a fresh worker is created), raise `workerTimeoutMs` via `configureWasm`, or split the input. Other worker failures (script load failure, un-cloneable payload, WASM init error inside the worker) keep the existing main-thread retry.
