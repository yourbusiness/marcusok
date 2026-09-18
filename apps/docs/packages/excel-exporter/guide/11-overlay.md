# Progress Overlay

An **optional** full-screen overlay with a progress bar, shown while an export runs. It lives behind its own subpath, so the main entry (`@marcusok/excel-exporter`) stays free of DOM code and callers who don't use it ship no extra bytes.

```ts
import { exportExcelWithOverlay } from "@marcusok/excel-exporter/overlay";

const result = await exportExcelWithOverlay({
  filename: "sales-2026",
  sheets: [{ name: "Sales", columns, data }],
});
```

The overlay appears after a short delay, blocks page interaction, and is removed when the export settles — on success **and** on failure.

## Options

```ts
await exportExcelWithOverlay(options, {
  delayMs: 200, // don't show at all if the export finishes within 200ms
  minVisibleMs: 300, // once shown, stay at least this long (no flash)
  fadeOutMs: 150,
  zIndex: 2147483000,
  container: document.body,
  blockInteraction: true, // false = visual cover only
  theme: "auto", // "auto" | "light" | "dark"
  text: {
    title: "正在导出 Excel",
    preparing: "准备中…",
    building: "正在构建工作簿…",
    downloading: "正在下载…",
    finishing: "即将完成…",
    hint: "数据量较大时可能需要数十秒，请勿关闭页面",
  },
});
```

| Option             | Default         | Description                                                                                       |
| ------------------ | --------------- | ------------------------------------------------------------------------------------------------- |
| `delayMs`          | `200`           | Delay before the overlay is mounted. An export that finishes sooner never shows it at all.        |
| `minVisibleMs`     | `300`           | Once shown, the overlay stays at least this long — it is removed on a delay rather than flashing. |
| `fadeOutMs`        | `150`           | Fade-out duration before the node is detached.                                                    |
| `zIndex`           | `2147483000`    | Overlay stacking level.                                                                           |
| `container`        | `document.body` | Mount target.                                                                                     |
| `blockInteraction` | `true`          | Blocks pointer and scroll events on the overlay. `false` leaves the page usable underneath.       |
| `theme`            | `"auto"`        | `"auto"` resolves via `prefers-color-scheme` at mount time.                                       |
| `text`             | Chinese         | Label overrides; the progress bar tracks the `hint` shown while indeterminate.                    |

`text.hint` is only rendered while the bar is indeterminate — a determinate bar shows its percentage instead.

## Indeterminate vs determinate

Progress comes from the library's existing callbacks, so the bar is only as good as the data behind it:

| Route                | Condition                  | Intermediate progress |
| -------------------- | -------------------------- | --------------------- |
| main + Workbook      | browser < 20,000 rows      | none                  |
| main + Fast stream   | Node, or explicit `stream` | every 1,000 rows      |
| Worker + Workbook    | auto, 20,000–49,999 rows   | none                  |
| Worker + Fast stream | auto, ≥ 50,000 rows        | every 1,000 rows      |

Only the Fast stream path emits values between `0` and `1` (`onProgress`). The overlay therefore renders an **animated sweep** until the first intermediate value arrives, then switches to a determinate bar. Workbook routes stay indeterminate for their whole run — that is the granularity of the data source, not a rendering problem.

The trailing `onProgress(1)` never closes the overlay and never fabricates a completed bar on a route that had no real progress; closing is driven by the promise settling.

## Blocking the main thread

`WorkbookBuilder.addSheet` and the Fast stream writer are **synchronous**. While they run the browser cannot repaint, which constrains how the overlay can behave:

- With the default `delayMs: 200`, if a blocking span is already in progress when the delay expires, the reveal is starved: the overlay **never appears** for that export. A late timer is explicitly cleared on close, so it can never pop up _after_ the export finished either.
- With `delayMs: 0` the overlay is mounted synchronously _before_ `exportExcel` is called, and a two-frame yield lets the browser paint it first. This is the only way to see the overlay on a blocking route — at the cost of a brief flash on fast exports.

Either way, the sweep animation keeps running during a block (it is a CSS transform animation, driven by the compositor), while the percentage and label freeze until the thread is free again.

## Handle form

For `exportTable`, `exportEcharts`, or a custom flow, drive the overlay yourself:

```ts
import { showExportOverlay } from "@marcusok/excel-exporter/overlay";

const overlay = showExportOverlay({ delayMs: 200 });
try {
  return await exportTable({
    columns,
    data,
    filename: "report",
    onProgress: (p) => overlay.handleProgress(p),
    onPhase: (phase, ms) => overlay.handlePhase(phase),
  });
} finally {
  overlay.close(); // idempotent
}
```

`exportExcelWithOverlay` is exactly this pattern: it **appends** to your existing `onProgress` / `onPhase` instead of replacing them, so a metrics panel wired to the same callbacks keeps working.

## Concurrency

Overlays share one DOM node, reference-counted: concurrent exports render into the same overlay (last writer wins) and it is removed when the last one closes. `exportTable`/`exportExcel` are safe to call concurrently; nothing leaks between runs.

## Node and SSR

Without a `document`, `showExportOverlay` returns a no-op handle and `exportExcelWithOverlay` exports normally — the same call site works in both environments, no branching required.

## Accessibility

The bar carries `role="progressbar"` (with `aria-valuenow` only while determinate), the label is an `aria-live="polite"` status region, and the container sets `aria-busy`. `prefers-reduced-motion: reduce` disables the sweep animation and all transitions.

Note that blocking is implemented at the pointer level, and the overlay deliberately does **not** apply `inert` to the page beneath it — keyboard users can still tab to elements that are visually covered. Applying `inert` to every sibling node is invasive enough to risk breaking host-page behaviour, so it is left out.

## See also

- [Progress and phase callbacks](./10-advanced#progress-and-phase-callbacks) — the underlying `onProgress` / `onPhase` contract.
- [Worker and stream modes](./06-worker-stream) — which route a given row count takes.
