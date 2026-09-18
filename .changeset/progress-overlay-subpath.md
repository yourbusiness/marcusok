---
"@marcusok/excel-exporter": minor
---

New optional subpath `@marcusok/excel-exporter/overlay`: a full-screen progress overlay for exports.

```ts
import { exportExcelWithOverlay } from "@marcusok/excel-exporter/overlay";

const result = await exportExcelWithOverlay(
  { filename: "report", sheets: [...] },
  { delayMs: 200 },
);
```

- **`exportExcelWithOverlay(options, overlayOptions?)`** wraps `exportExcel`, showing a blocking overlay while the export runs and removing it when the promise settles — on success and on failure alike. It **appends to** your existing `onProgress` / `onPhase` callbacks rather than replacing them, so any UI already wired to those keeps working.
- **`showExportOverlay(overlayOptions?)`** returns a handle (`handleProgress` / `handlePhase` / `close`) for `exportTable`, `exportEcharts`, and custom flows.
- Options: `delayMs` (default 200 — a faster export never shows the overlay at all), `minVisibleMs` (default 300), `fadeOutMs`, `zIndex`, `container`, `blockInteraction` (default true), `theme` (`auto` / `light` / `dark`), and `text` label overrides.
- The overlay is **indeterminate until real progress arrives**: only the Fast stream routes emit `onProgress` values between 0 and 1. Workbook routes (browser < 20,000 rows, Worker 20,000–49,999 rows) show the animated sweep for their whole run. The trailing `onProgress(1)` never closes the overlay and never fabricates a completed bar on a route that had no real progress.
- `delayMs` also decouples the overlay from main-thread blocking: on a route whose build blocks the thread, a reveal that could not fire before the block **never happens** rather than appearing after the export finished. Use `delayMs: 0` to mount before the export starts and let a two-frame yield paint it first.
- No `document` (Node/SSR) returns a no-op handle, so the same call site works in both environments. No new runtime dependencies.

The main entry is untouched — `src/index.ts` and `src/types.ts` are unchanged, and the overlay is a separate build entry, so callers who don't import it ship exactly the same bytes as before.

Also in this release: a new `Progress Overlay` guide page (English and Chinese), an overlay toggle in the play and docs-site demos, and `overlay-state.test.ts` / `overlay.test.ts` (35 new tests; `happy-dom` added as a test-only devDependency).
