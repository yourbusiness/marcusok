---
"@marcusok/excel-preview": patch
---

Fix merged cells disappearing from frozen bands: the frozen row/column bands follow the main viewport on their unpinned axis, so a wide merge anchored inside the band (e.g. an `A1:H1` title in a frozen first row) vanished once scrolling moved its anchor out of the window. Band layers now recall off-window merge anchors the same way the main quadrant does.
