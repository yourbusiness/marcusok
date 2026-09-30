# Usage

## The handle

`showProgressOverlay(options)` returns a handle with exactly three methods, all exception-safe and safe to call after `close()`:

- `setProgress(p)` — feed progress values. Any value in `(0, 1)` switches the overlay to the determinate bar. The leading `0` is ignored; the trailing `1` only completes an existing determinate bar — a task that never reported intermediate progress stays on the spinner (that is the data source's granularity, not a rendering problem).
- `setPhase(key)` — switch the label via the `text.phases` map.
- `close()` — idempotent. Call it in a `finally`; never key it off `setProgress(1)`, which also fires on failed runs.

## Options

| Option             | Default         | Description                                                                                                                                      |
| ------------------ | --------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `delayMs`          | `200`           | Delay before mounting. A task that finishes sooner never shows the overlay at all.                                                               |
| `minVisibleMs`     | `300`           | Once shown, stay at least this long — removed on a delay rather than flashing.                                                                   |
| `fadeOutMs`        | `150`           | Fade-out duration before the node is detached.                                                                                                   |
| `zIndex`           | `2147483000`    | Overlay stacking level.                                                                                                                          |
| `container`        | `document.body` | Mount target.                                                                                                                                    |
| `blockInteraction` | `true`          | Blocks pointer and scroll events on the overlay itself (the host page's `overflow` is never touched). `false` leaves the page usable underneath. |
| `theme`            | `"auto"`        | `"auto"` resolves via `prefers-color-scheme` at mount time.                                                                                      |
| `text.title`       | `"请稍候"`      | Panel title.                                                                                                                                     |
| `text.initial`     | `"正在处理…"`   | Label before the first `setPhase`.                                                                                                               |
| `text.phases`      | `{}`            | `key → label` map; `setPhase(key)` looks the label up here. An uncovered key renders as-is.                                                      |
| `text.hint`        | (see above)     | Extra line under the label, indeterminate mode only.                                                                                             |

## Blocking the main thread

A long **synchronous** span in your task (building a big object in one go, sync XHR-style processing) constrains what the overlay can do, because the browser cannot repaint mid-task:

- With the default `delayMs: 200`, if the blocking span is already in progress when the delay expires, the reveal is starved: the overlay **never appears** for that task. The late timer is explicitly cleared on close, so it can never pop up _after_ the task ends either.
- With `delayMs: 0` the overlay mounts synchronously _before_ your task starts. Await `nextPaint()` (exported from this package) between `showProgressOverlay` and the blocking call — that yield is the only guarantee the browser ever paints the overlay, at the cost of a brief flash on fast tasks.

Either way, animations keep running during a block (CSS transform animations are driven by the compositor), while the percentage and label freeze until the thread is free.

```ts
const overlay = showProgressOverlay({ delayMs: 0 });
try {
  await nextPaint(); // let the browser paint the overlay first
  heavySyncWork(); // blocking span — overlay is visible & animating
} finally {
  overlay.close();
}
```

`nextPaint()` yields two animation frames, with a 250ms timer racing them: in hidden tabs (where rAF is paused indefinitely) the timer wins so a background task can never hang on the yield.

## Concurrency

Overlays share one DOM node, reference-counted: concurrent tasks render into the same overlay (last writer wins) and it is removed when the last one closes. Two consequences of the shared node:

- While the overlay is still pending, the reveal waits until every concurrent caller's `delayMs` has elapsed — a second task with a longer delay is never revealed early by the first task's timer.
- A second task started while the overlay is already visible takes over the display immediately at `show` time (title/hint/label reset to the new caller's texts), not on its first progress/phase event.

If an early-quitting caller's pending reveal timer fires after it closed, the texts shown belong to the most recent still-active caller.

## Node and SSR

Without a `document`, `showProgressOverlay` returns a no-op handle — the same call site works in both environments, no branching required.

## Accessibility

The bar carries `role="progressbar"` (with `aria-valuenow` only while determinate), the label is an `aria-live="polite"` status region, the spinner is `aria-hidden` decoration, and the container sets `aria-busy`. The percentage is a separate node outside the live region, so screen readers do not re-announce every progress tick.

Blocking is pointer-level only; the overlay deliberately does **not** apply `inert` to the page beneath it — keyboard users can still tab to visually covered elements. Applying `inert` to every sibling node is invasive enough to risk breaking host-page behaviour, so it is left out.

## Using it from @marcusok packages

`@marcusok/excel-exporter` 2.8+ enables this overlay by default (`overlay: false` to opt out, `overlay: {…}` to customize) and contributes export-flavored texts. If you drive the handle yourself for a custom flow, re-exported bindings are available from `@marcusok/excel-exporter/overlay` (legacy subpath) — or import this package directly and keep stage keys of your own.
