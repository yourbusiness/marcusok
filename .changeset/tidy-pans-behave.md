---
"@marcusok/excel-preview": minor
---

Add parse cancellation. `createPreview().destroy()` now aborts an in-flight parse immediately instead of letting it run to completion: the boot path returns silently (`onError` is never fired for a cancellation), the main-thread fallback path gives up at the next `await` boundary, and a late worker result is dropped via its already-cleared pending slot (the shared worker is not terminated — that would reject sibling requests). The low-level `parseWorkbookBytes(bytes, opts)` accepts an optional `AbortSignal` via `opts.signal` and rejects with an `ABORTED`-coded error when aborted. `ABORTED` is a new `PreviewErrorCode` member and is never routed to the main-thread fallback.
