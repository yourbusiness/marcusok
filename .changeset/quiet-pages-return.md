---
"@marcusok/pdf-preview": patch
"@marcusok/ppt-preview": patch
---

Fix lazy-render re-entry races that could leave a page permanently blank after scrolling away during an in-flight render and quickly scrolling back: visibility changes are now tracked while a page is rendering, cancelled PDF renders are re-queued when the page is visible again, and completed PPT renders are no longer released by a stale release request. Also clear stale PDF release state after failed renders, silence late callbacks after document-level failures, release the PPTX engine reference on load-time destroy, and normalize synchronous `load()` failures into `onError` instead of an unhandled rejection.
