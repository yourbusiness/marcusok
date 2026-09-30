---
"@marcusok/xlsx-core": patch
---

Correct the stale entry-header comment (modern-xlsx is a declared runtime-free dependency kept for d.ts resolution, not a devDependency) and document that `workerTimeoutMs` must be positive — 0/negative values time out immediately rather than disabling the timeout.
