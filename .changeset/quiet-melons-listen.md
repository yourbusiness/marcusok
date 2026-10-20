---
"@marcusok/excel-preview": patch
---

Cap the total number of merge-covered cells expanded during layout at 2,000,000 cells across all merges, on top of the existing per-merge 1,000,000 cap. Files declaring many near-limit merges could previously stack the main-thread expansion into multi-second freezes (10 × 1M merges ≈ 5.5s measured); over-budget merges now degrade exactly like single oversized ones — the anchor index is kept so the merge still renders with its span, while covered cells fall back to rendering their own content.
