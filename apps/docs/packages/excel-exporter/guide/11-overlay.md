# Progress Overlay

Since 2.8.0 every export shows an **optional-by-configuration** full-screen overlay by default — no extra import, no wrapper function. Set `overlay: false` to turn it off entirely; pass an options object to customize it. The overlay itself lives in a separate shared package, [@marcusok/progress-overlay](/packages/progress-overlay/); this guide covers how the exporter drives it.

```ts
import { exportExcel } from "@marcusok/excel-exporter";

const result = await exportExcel({
  filename: "sales-2026",
  sheets: [{ name: "Sales", columns, data }],
  // overlay: false,            // <- opt out entirely
  // overlay: { delayMs: 0 },   // <- customize
});
```

The overlay appears after a short delay, blocks page interaction, and is removed when the export settles — on success **and** on failure. In Node/SSR it is a no-op (there is no `document`).

`exportTable` and `exportEcharts` delegate to `exportExcel`, but their own option types (`TableExportOptions` / `EChartsExportOptions`) do **not** carry `overlay` and their converters drop the field — passing it there is a TypeScript error and is silently ignored at runtime, so the default overlay still shows. To control the overlay while starting from those data shapes, convert first with `tableToSheet` / `echartsToSheet` and call `exportExcel` yourself (see [Table and ECharts entries](/packages/excel-exporter/api/05-table-and-echarts)).

## Option values

| Value                    | Behaviour                                                                                                                                                                                                                                                                                                |
| ------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| omitted / `true`         | Overlay with the default export texts (default since 2.8.0).                                                                                                                                                                                                                                             |
| `false`                  | No overlay at all: nothing is mounted, no extra frame is yielded, callbacks run exactly as before the feature existed.                                                                                                                                                                                   |
| `ProgressOverlayOptions` | Customize: texts, `delayMs`, theme, `blockInteraction`, … (see the [shared package](/packages/progress-overlay/guide/01-usage)). Custom texts are **merged** with the export defaults — overriding `text.title` keeps the built-in stage labels, so phase keys like `building` never render as raw keys. |

```ts
await exportExcel({
  filename: "sales-2026",
  sheets: [{ name: "Sales", columns, data }],
  // `overlay` is a field of ExportOptions — every ProgressOverlayOptions
  // entry is customizable there. The values below are the built-in export
  // defaults; pass only what you want to change.
  overlay: {
    delayMs: 200, // don't show at all if the export finishes within 200ms
    minVisibleMs: 300, // once shown, stay at least this long (no flash)
    fadeOutMs: 150,
    zIndex: 2147483000,
    container: document.body,
    blockInteraction: true, // false = visual cover only
    theme: "auto", // "auto" | "light" | "dark"
    text: {
      title: "正在导出 Excel",
      initial: "准备中…",
      phases: {
        building: "正在构建工作簿…",
        downloading: "正在下载…",
        finishing: "即将完成…",
      },
      hint: "数据量较大时可能需要数十秒，请勿关闭页面",
    },
  },
});
```

`text.hint` is only rendered while the bar is indeterminate — a determinate bar shows its percentage instead.

## Indeterminate vs determinate

Progress comes from the library's existing callbacks (`onProgress` / `onPhase` are **chained**, never replaced — a metrics panel wired to the same callbacks keeps working), so the overlay is only as good as the data behind it:

| Route                | Condition                                                            | Intermediate progress |
| -------------------- | -------------------------------------------------------------------- | --------------------- |
| main + Workbook      | browser < 20,000 rows                                                | none                  |
| main + Fast stream   | Node (auto ≥ 50,000 rows, or explicit `stream` — Node has no Worker) | every 1,000 rows      |
| Worker + Workbook    | auto, 20,000–49,999 rows                                             | none                  |
| Worker + Fast stream | auto ≥ 50,000 rows, or explicit `stream` in the browser              | every 1,000 rows      |

Only the Fast stream path emits values between `0` and `1` (`onProgress`). The overlay therefore renders an **animated spinner** until the first intermediate value arrives, then switches to a determinate bar. Workbook routes stay indeterminate for their whole run — that is the granularity of the data source, not a rendering problem.

The trailing `onProgress(1)` never closes the overlay and never fabricates a completed bar on a route that had no real progress; closing is driven by the promise settling.

## Blocking the main thread

`WorkbookBuilder.addSheet` and the Fast stream writer are **synchronous**. While they run the browser cannot repaint, which constrains how the overlay can behave:

- With the default `delayMs: 200`, if a blocking span is already in progress when the delay expires, the reveal is starved: the overlay **never appears** for that export. This is why a small, warm (already-WASM-initialized) export on the main route shows no overlay at all — the build starts before the delay expires and finishes under it. A late timer is explicitly cleared on close, so it can never pop up _after_ the export finished either.
- With `delayMs: 0` the overlay is mounted synchronously _before_ `exportExcel` runs its build, and a two-frame yield (`nextPaint`) lets the browser paint it first. This is the only way to see the overlay on a blocking route — at the cost of a brief flash on fast exports.

Either way, the spinner keeps spinning during a block (it is a CSS transform animation, driven by the compositor), while the percentage and label freeze until the thread is free again.

## Driving the overlay yourself

For a custom flow around the low-level entry points, drive [@marcusok/progress-overlay](/packages/progress-overlay/) directly — the exporter's overlay option is exactly this protocol with export-flavored texts:

```ts
import { showProgressOverlay } from "@marcusok/progress-overlay";

const overlay = showProgressOverlay({
  text: { title: "正在导出 Excel", phases: { building: "正在构建工作簿…" } },
});
try {
  // exportAsStream's second argument is the callback itself, not an options object
  return await exportAsStream(sheets, (p) => overlay.setProgress(p));
} finally {
  overlay.close(); // idempotent
}
```

## Legacy subpath

`@marcusok/excel-exporter/overlay` (pre-2.8) still exists for compatibility: `exportExcelWithOverlay(options, overlay?)` is now a thin wrapper equivalent to `exportExcel({ ...options, overlay })`, and `showExportOverlay` forwards to the shared package. Note the text shape changed with the move: the old flat fields (`text.building`, `text.downloading`, …) are now a `text.phases` map, and the handle methods are `setProgress` / `setPhase(key)` instead of `handleProgress` / `handlePhase`.

## Concurrency

Overlays share one DOM node, reference-counted: concurrent exports render into the same overlay (last writer wins) and it is removed when the last one closes. `exportTable`/`exportExcel` are safe to call concurrently; nothing leaks between runs. The rendered content (title, hint, label, progress) always belongs to the most recent `show` / progress / phase event's caller. See the [shared package's concurrency notes](/packages/progress-overlay/guide/01-usage#concurrency) for the full semantics.

## Node and SSR

Without a `document`, the overlay is a no-op handle and `exportExcel` exports normally — the same call site works in both environments, no branching required.

## Accessibility

The bar carries `role="progressbar"` (with `aria-valuenow` only while determinate), the label is an `aria-live="polite"` status region, and the container sets `aria-busy`. `prefers-reduced-motion: reduce` disables the spinner animation and all transitions.

Note that blocking is implemented at the pointer level, and the overlay deliberately does **not** apply `inert` to the page beneath it — keyboard users can still tab to elements that are visually covered. Applying `inert` to every sibling node is invasive enough to risk breaking host-page behaviour, so it is left out.

## See also

- [@marcusok/progress-overlay](/packages/progress-overlay/) — the shared overlay package this option drives (options table, blocking-thread trade-offs, concurrency semantics).
- [Progress and phase callbacks](./10-advanced#progress-and-phase-callbacks) — the underlying `onProgress` / `onPhase` contract.
- [Worker and stream modes](./06-worker-stream) — which route a given row count takes.
