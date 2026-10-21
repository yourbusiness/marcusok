---
"@marcusok/excel-exporter": minor
---

Add `signal?: AbortSignal` to export options (including `exportTable`/`exportEcharts`) and rework the shared export worker queue. Requests are now serialized through a main-thread FIFO queue, a timeout starts when a request is dispatched (queue wait excluded), and a timed-out/aborted request no longer rejects unrelated queued exports — they continue on a fresh worker. Aborting resolves a structured `AbortError` and skips the degradation chain.
