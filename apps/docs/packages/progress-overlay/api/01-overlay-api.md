# API: showProgressOverlay & Customization

Reference for the two exports and the visual surface. For the concepts behind them — when the overlay appears, how the handle behaves under concurrency, the blocking-thread trade-offs — start with the [usage guide](../guide/01-usage).

## showProgressOverlay

```ts
showProgressOverlay(options?: ProgressOverlayOptions): ProgressOverlayHandle
```

Creates the overlay (or joins the one concurrent tasks already share) and returns a handle. Without a `document` — Node/SSR — or when `container` resolves to `null`, it returns a no-op handle instead, so a single call site works in both environments.

Styles are injected once, on first use, as a `<style id="mxe-overlay-style">` in `document.head`; every later call reuses them.

## nextPaint

```ts
nextPaint(): Promise<void>
```

Yields two animation frames, raced against a 250 ms timer — in a hidden tab, where `requestAnimationFrame` is paused indefinitely, the timer wins so a background task can never hang on the yield. This is the only way to guarantee the browser paints the overlay before a long synchronous span; see [Blocking the main thread](../guide/01-usage#blocking-the-main-thread).

## ProgressOverlayOptions

| Option             | Type                          | Default                                | Description                                                                                                                                     |
| ------------------ | ----------------------------- | -------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `delayMs`          | `number`                      | `200`                                  | Delay before mounting. A task that finishes sooner never shows the overlay at all. Negative values are clamped to `0`                           |
| `minVisibleMs`     | `number`                      | `300`                                  | Once shown, stay at least this long — removed on a delay rather than flashing. Negative values are clamped to `0`                               |
| `fadeOutMs`        | `number`                      | `150`                                  | Fade-out duration before the node is detached. Negative values are clamped to `0`                                                               |
| `zIndex`           | `number`                      | `2147483000`                           | Overlay stacking level                                                                                                                          |
| `container`        | `HTMLElement`                 | `document.body`                        | Mount target                                                                                                                                    |
| `blockInteraction` | `boolean`                     | `true`                                 | Blocks pointer and scroll events on the overlay itself (the host page's `overflow` is never touched). `false` leaves the page usable underneath |
| `theme`            | `"auto" \| "light" \| "dark"` | `"auto"`                               | `"auto"` resolves via `prefers-color-scheme` at mount time                                                                                      |
| `text.title`       | `string`                      | `"请稍候"`                             | Panel title                                                                                                                                     |
| `text.initial`     | `string`                      | `"正在处理…"`                          | Label before the first `setPhase`                                                                                                               |
| `text.phases`      | `Record<string, string>`      | `{}`                                   | `key → label` map; `setPhase(key)` looks the label up here. An uncovered key renders as-is                                                      |
| `text.hint`        | `string`                      | `"任务可能需要一些时间，请勿关闭页面"` | Extra line under the label, indeterminate mode only                                                                                             |

### Where an option applies

`container`, `theme`, `zIndex` and `blockInteraction` are properties of the **DOM node**, not of the call: while concurrent tasks share one node, they are taken from whichever call created it. A later caller passing different values has them silently ignored until the overlay is fully torn down and rebuilt (which happens when a call arrives with a different `container`). Text options (`text.*`) are per-caller and merge into the display as described in [Concurrency](../guide/01-usage#concurrency).

## ProgressOverlayHandle

| Method           | Contract                                                                                                                                                                                                                                                                                                                    |
| ---------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `setProgress(p)` | Feeds progress. Any value in `(0, 1)` switches the overlay to the determinate bar. The leading `0` is ignored; the trailing `1` only completes a bar that is already determinate — a task that never reported intermediate progress stays on the spinner. Values are **not** monotonic: a lower value redraws a shorter bar |
| `setPhase(key)`  | Switches the label via the `text.phases` map. An uncovered key renders as-is, so typos surface immediately. Repeating the same key is ignored (it would re-announce on screen readers)                                                                                                                                      |
| `close()`        | Idempotent, and still callable after closing. Call it in a `finally` — never key it off `setProgress(1)`, which also fires on failed runs                                                                                                                                                                                   |

All three are exception-safe: every method body is wrapped, because the overlay is decoration and must never break the task it decorates.

Two further behaviours are worth knowing when you wire progress in:

- **Values arriving before the reveal are kept.** If the overlay has not been revealed yet (it is still inside `delayMs`), `setProgress` / `setPhase` update its state but skip rendering — the reveal then shows those values immediately instead of an empty panel.
- **A handle can go stale.** If another call mounts the overlay into a different `container`, the previous node is dropped and any handle pointing at it becomes a silent no-op (no throw, no warning). The usual cause is mixing `container` values across concurrent tasks — pick one.

## Customizing the look

### CSS variables

All nine are declared on `.mxe-overlay` for the light theme and re-declared under `.mxe-overlay[data-mxe-theme="dark"]`:

| Variable             | Light                      | Dark                       |
| -------------------- | -------------------------- | -------------------------- |
| `--mxe-backdrop`     | `rgba(82, 82, 91, .32)`    | `rgba(0, 0, 0, .55)`       |
| `--mxe-panel-bg`     | `rgba(255, 255, 255, .72)` | `rgba(24, 24, 27, .65)`    |
| `--mxe-panel-border` | `rgba(0, 0, 0, .06)`       | `rgba(255, 255, 255, .08)` |
| `--mxe-panel-shadow` | layered drop shadow        | layered drop shadow        |
| `--mxe-title`        | `#18181b`                  | `#fafafa`                  |
| `--mxe-text`         | `#71717a`                  | `#a1a1aa`                  |
| `--mxe-hint`         | `#a1a1aa`                  | `#71717a`                  |
| `--mxe-track`        | `rgba(0, 0, 0, .08)`       | `rgba(255, 255, 255, .14)` |
| `--mxe-fill`         | `#171717`                  | `#fafafa`                  |

Because the stylesheet is appended to `document.head` on first use, it lands **after** the stylesheets your app loaded at startup — so an equally specific override in your own CSS loses. Raise specificity instead of relying on order:

```css
/* wins over the injected `.mxe-overlay { --mxe-fill: #171717 }` */
html .mxe-overlay {
  --mxe-fill: #2563eb;
  --mxe-track: rgba(37, 99, 235, 0.18);
}
```

### Class names and data attributes

| Selector / attribute                                     | Element                                                                  |
| -------------------------------------------------------- | ------------------------------------------------------------------------ |
| `.mxe-overlay`                                           | Root: backdrop, blur, `aria-busy`, and the two data attributes below     |
| `.mxe-overlay[data-mxe-theme]`                           | `"light"` / `"dark"` — resolved theme                                    |
| `.mxe-overlay[data-mxe-mode]`                            | `"indeterminate"` / `"determinate"` — drives which half is visible       |
| `.mxe-panel`                                             | Glass panel (title, bar row, hint)                                       |
| `.mxe-spinner`, `.mxe-spinner-track`, `.mxe-spinner-arc` | The indeterminate SVG spinner and its two circles                        |
| `.mxe-title`                                             | Panel title                                                              |
| `.mxe-bar`, `.mxe-fill`                                  | Progress track and its fill                                              |
| `.mxe-row`, `.mxe-label`, `.mxe-percent`                 | Label/percentage row; the percentage sits outside the `aria-live` region |
| `.mxe-hint`                                              | Hint line (indeterminate only)                                           |

Mode and theme are the mechanism behind the two-state behaviour: `data-mxe-mode` is what hides the bar and the percentage (or the spinner and the hint), so a custom stylesheet can restyle around them without touching the render logic.

## Used by

`@marcusok/excel-exporter` 2.8+ drives this protocol for you through its [`overlay` option](/packages/excel-exporter/guide/11-overlay) and contributes export-flavoured texts. Import this package directly when you want the overlay for a non-export task, or when you are driving the low-level export entries yourself.
