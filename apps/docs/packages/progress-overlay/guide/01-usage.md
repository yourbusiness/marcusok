# Usage

## The handle

`showProgressOverlay(options)` returns a handle with exactly three methods — `setProgress(p)`, `setPhase(key)` and `close()`. All three are exception-safe and still callable after `close()`. The [API reference](../api/01-overlay-api#progressoverlayhandle) specifies each one's contract: what the leading `0` and the trailing `1` do, why a repeated phase key is ignored, and when a handle goes stale.

## Options

Every option, with its type, default and scope, is tabulated in the [API reference](../api/01-overlay-api#progressoverlayoptions). The timing defaults are `delayMs: 200`, `minVisibleMs: 300` and `fadeOutMs: 150`.

That page also covers [customizing the look](../api/01-overlay-api#customizing-the-look) through CSS variables and class names — including the `data-mxe-mode` attribute that drives the two-state behaviour described below.

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
