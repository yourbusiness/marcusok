---
"@marcusok/excel-preview": patch
---

Docs: describe the main-thread fallback precisely. The README now states both trigger conditions (no `Worker` or no `window` global) and that a worker which fails to construct/load falls back permanently for the session, instead of attributing the fallback to "no Worker global" (Node/SSR) alone.
