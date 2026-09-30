# @marcusok/progress-overlay

A framework-agnostic full-screen progress overlay: an animated spinner while task length is unknown, a slim percentage bar once real progress arrives. Glass panel, light/dark themes, zero runtime dependencies — the shared UI layer behind [@marcusok/excel-exporter](/packages/excel-exporter/)'s `overlay` option, usable standalone for any long-running task (report generation, batch uploads, imports…).

## Install

```bash
pnpm add @marcusok/progress-overlay
```

No runtime dependencies at all — pure DOM plus a one-time injected stylesheet.

## Quick example

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

## Capabilities

| Capability         | Description                                                                                                                                                                                                                |
| ------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Two-state adaptive | Spinner + hint while duration is unknown; percentage bar as soon as a `(0, 1)` progress value arrives. No fabrication: a trailing `1` alone never fabricates a completed bar.                                              |
| Neutral glass look | Zinc-grey palette that does not clash with host branding; blurred backdrop, glass panel with inner highlight, `prefers-color-scheme` themes (`auto` / `light` / `dark`).                                                   |
| Delay gating       | `delayMs: 200` by default — tasks that finish sooner never flash the overlay at all; a minimum visible time avoids one-frame blinks.                                                                                       |
| Concurrency        | One reference-counted DOM node shared by concurrent tasks (last writer wins); texts always belong to the most recent active caller.                                                                                        |
| Robustness         | Every handle method is exception-safe (the overlay must never break the task it decorates); Node/SSR returns a no-op handle — same call site, no branching.                                                                |
| Accessibility      | `role="progressbar"` with `aria-valuenow` only while determinate; the label is an `aria-live="polite"` region and the percentage is kept out of it (no re-announcing per tick); `prefers-reduced-motion` stops all motion. |

## Where the texts come from

The package has no business vocabulary: default texts are generic (`请稍候` / `正在处理…`), stage labels are a caller-defined `key → text` map (`text.phases`), and an uncovered key renders as-is so typos surface immediately. `@marcusok/excel-exporter` maps its `ExportPhase` events onto this protocol and contributes the "正在导出 Excel" texts — see the [exporter's overlay guide](/packages/excel-exporter/guide/11-overlay).

Continue with the [usage guide](./guide/01-usage) for the full option table, the blocking-thread trade-offs (`delayMs: 0` + `nextPaint()`), and concurrency semantics.
