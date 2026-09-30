# @marcusok/progress-overlay

A framework-agnostic full-screen progress overlay: an animated spinner while work length is unknown, a slim percentage bar once real progress arrives. Glass panel, light/dark themes, zero dependencies, zero framework code — the shared UI layer behind `@marcusok/excel-exporter` (enabled there by default), usable standalone for any long-running task.

```ts
import { showProgressOverlay } from "@marcusok/progress-overlay";

const overlay = showProgressOverlay({
  text: {
    title: "Generating report",
    phases: { building: "Building workbook…", downloading: "Downloading…" },
  },
});

try {
  await runTask({
    onProgress: (p) => overlay.setProgress(p), // 0..1
    onStage: (key) => overlay.setPhase(key), // key from text.phases
  });
} finally {
  overlay.close(); // idempotent; the only correct close site
}
```

## How it renders

- **Indeterminate** (before any intermediate progress): spinner → title → hint, vertically centered on a neutral dimmed backdrop with a blur.
- **Determinate** (as soon as a progress value in `(0, 1)` arrives): title, a 4px rounded bar, and a label row with the percentage right-aligned. The hint disappears.
- Light and dark themes via `prefers-color-scheme` (or forced); colors are neutral zinc greys so the overlay does not clash with host-site branding.
- `prefers-reduced-motion: reduce` stops the spinner and every transition.

Progress granularity is the caller's: only callers that emit intermediate values ever reach the determinate mode. A task that only reports `0` and `1` stays on the spinner for its whole run — that is the data source, not a rendering problem.

## API

### `showProgressOverlay(options?): ProgressOverlayHandle`

Returns a handle whose every method is exception-safe: the overlay is an enhancement and must never break the task it decorates. In Node/SSR (no `document`) the handle is a no-op — the same call site works everywhere, no branching.

| Option             | Default         | Description                                                                                 |
| ------------------ | --------------- | ------------------------------------------------------------------------------------------- |
| `delayMs`          | `200`           | Delay before mounting. A task that finishes sooner never shows the overlay at all.          |
| `minVisibleMs`     | `300`           | Once shown, stay at least this long — removed on a delay instead of flashing.               |
| `fadeOutMs`        | `150`           | Fade-out duration before the node is detached.                                              |
| `zIndex`           | `2147483000`    | Overlay stacking level.                                                                     |
| `container`        | `document.body` | Mount target.                                                                               |
| `blockInteraction` | `true`          | Blocks pointer and scroll events on the overlay. `false` leaves the page usable underneath. |
| `theme`            | `"auto"`        | `"auto"` resolves via `prefers-color-scheme` at mount time.                                 |
| `text.title`       | `"请稍候"`      | Panel title.                                                                                |
| `text.initial`     | `"正在处理…"`   | Label before the first `setPhase`.                                                          |
| `text.phases`      | `{}`            | `key → label` map; `setPhase(key)` looks the label up here. An uncovered key renders as-is. |
| `text.hint`        | (see above)     | Extra line under the label, indeterminate mode only.                                        |

### Handle

- `setProgress(p)` — feed `onProgress` values. Any value in `(0, 1)` switches to the determinate bar; the leading `0` is ignored; the trailing `1` only completes an existing determinate bar (it never fabricates one).
- `setPhase(key)` — switch the label via `text.phases`.
- `close()` — idempotent. Call it in a `finally`; never key it off `setProgress(1)`, which also fires on failed runs.

### `nextPaint(): Promise<void>`

Yields two animation frames (with a 250ms timer fallback for hidden tabs, where rAF is paused). Await it after `showProgressOverlay({ delayMs: 0 })` and before a **synchronous** long task — otherwise the browser may never paint the overlay at all. With the default `delayMs: 200`, a blocking span that starts before the delay expires starves the reveal entirely and the overlay never appears (the late timer is cleared on close, so it also never pops up after the task ends).

## Concurrency

Overlays share one DOM node, reference-counted: concurrent tasks render into the same overlay (last writer wins) and it is removed when the last one closes. While the overlay is still pending, the reveal waits until every concurrent caller's `delayMs` has elapsed; a second task started while it is already visible takes over the display immediately at `show` time (title/hint/label reset to the new caller's text).

## Accessibility

The bar carries `role="progressbar"` (with `aria-valuenow` only while determinate; the percentage itself is kept out of the `aria-live` label so screen readers do not re-announce every progress tick), the label is an `aria-live="polite"` status region, the spinner is `aria-hidden` decoration, and the container sets `aria-busy`.

Blocking is pointer-level only; the overlay deliberately does not apply `inert` to the page beneath it.

## License

MIT
