---
"@marcusok/pdf-preview": patch
---

Align the text layer on rotated pages (`/Rotate` 90/180/270 transform rules), make `fit-page` measure `max-height` containers before placeholders exist and reserve the wrapper's vertical padding so the page no longer overflows, report a missing `IntersectionObserver` as `UNKNOWN` (environment) instead of `CORRUPT`, and fall back to the packaged asset directories when `workerSrc` is a non-hierarchical `blob:`/`data:` URL.
