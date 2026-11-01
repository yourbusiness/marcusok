---
"@marcusok/ppt-preview": patch
---

Ignore percentage `max-height` wrappers when detecting the scroll root (they are not height-constrained under an auto-height parent, which previously forced full-deck rendering and made navigation a silent no-op), and hand the engine the already-owned source buffer directly instead of copying the file a second time during load.
