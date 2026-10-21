---
"@marcusok/excel-exporter": patch
---

Validate and write merge ranges in linear time. Overlap detection, `addMergeCell` deduplication and repeated validation were quadratic: 40,000 disjoint merges took ~12.8s on the Workbook path and ~2.2s on the stream path. The new sweep-line validator plus batched merge insertion bring the same export down to ~0.26s/~0.11s, and avoid a user-input-driven denial-of-service surface.
