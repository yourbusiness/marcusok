---
"@marcusok/excel-preview": patch
---

Fix a ResizeObserver leak on repeated `render()` calls (e.g. loading another file into the same renderer instance). The previous observer — still watching the detached viewport — is now disconnected before a new one is created, so observers no longer accumulate until `destroy()`.
