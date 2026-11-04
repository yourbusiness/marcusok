---
"@marcusok/excel-preview": patch
---

`onParsed`'s `duration.parse` on the worker route now reports the pure parse time measured inside the worker (sniff → model build) instead of a main-thread wall clock that also covered FIFO queue wait — the two routes used the same field name with different meanings, and the worker-route number jittered with concurrent parses. The README also documents that the engine ships bundled with zero runtime dependencies.
