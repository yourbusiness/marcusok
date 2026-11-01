---
"@marcusok/excel-exporter": patch
---

Fix the bundled progress overlay so a concurrent task that used a long `delayMs` and closes early no longer keeps the shared overlay hidden until its original deadline; remaining callers now reveal at their own deadline.
