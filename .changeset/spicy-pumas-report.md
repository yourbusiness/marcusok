---
"@marcusok/ppt-preview": patch
---

Fix an unhandled rejection when post-load initialization fails (e.g. missing `IntersectionObserver` in very old WebViews, or SSR misuse): the error is now normalized and reported via `onError` with full cleanup instead of escaping as an unhandled promise rejection, matching the boot contract of the other preview packages.
