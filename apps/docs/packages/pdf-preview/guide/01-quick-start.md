# Quick Start

Requirement: a browser DOM. Importing the package is SSR-safe (the module top level touches no DOM globals — verified against pdfjs-dist 6.4.299), but `createPdfPreview` itself must run in the browser.

## 1. Install

```bash
pnpm add @marcusok/pdf-preview
```

Everything ships in the package: the pdf.js engine is bundled into `dist/index.js`, the worker into `dist/pdf.worker.min.mjs`, and the CJK / standard-font assets into `dist/cmaps/` + `dist/standard_fonts/` (~4.5 MB payload total). Zero transitive dependencies.

## 2. Mount a preview

```ts
import { createPdfPreview } from "@marcusok/pdf-preview";

const preview = createPdfPreview(container, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
  onRendered: ({ pageCount, firstPageMs, durationMs }) => {
    console.log(`${pageCount} pages; first page painted in ${firstPageMs}ms`);
  },
  onError: (e) => {
    // e.code: UNSUPPORTED | PASSWORD_PROTECTED | CORRUPT | UNKNOWN
    console.error(e.code, e.message);
  },
});
```

Pages render lazily as they enter the viewport — `onRendered` fires when the **first page** is painted (`pageCount` is the document total). There is no "all pages rendered" event in a lazy pipeline; that is by design.

## 3. Unmount

```ts
preview.destroy();
```

`destroy()` disconnects the render observers, cancels any in-flight render task, terminates the pdf.js worker and detaches the DOM synchronously. Late callbacks go silent (cancellation is not an error). Calling it twice is safe; the container can be reused immediately.

## Asset wiring — the one thing to get right

The worker is located by a `new URL("./pdf.worker.min.mjs", import.meta.url)` literal, which Vite/webpack asset pipelines pick up automatically (emitted as a base-aware hashed asset) — **zero config** for the worker under bundlers, and correct as-is for direct ESM/CDN consumption of the package `dist`.

`cmaps/` and `standard_fonts/` are **directories** — bundlers never copy those. What works per scenario:

| Scenario                                        | Worker                                          | cmaps / standard_fonts                                                                                         |
| ----------------------------------------------- | ----------------------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| Direct ESM / CDN import of the package          | resolved from `dist`                            | derived from the worker URL (same directory) — works as-is                                                     |
| Vite / webpack, defaults                        | emitted automatically ✅                        | derived URL points into the hashed-assets directory → **404, silent CJK degradation** ❌                       |
| Vite + `vite-static-copy` / webpack + copy plug | emitted automatically ✅                        | copy `dist/cmaps` + `dist/standard_fonts` to your public dir, pass explicit `cMapUrl`/`standardFontDataUrl` ✅ |
| Self-hosting via subpath exports                | `@marcusok/pdf-preview/dist/pdf.worker.min.mjs` | `@marcusok/pdf-preview/dist/cmaps/*` (directory wildcards are exported) ✅                                     |

The recommended bundler setup (Vite):

```ts
// vite.config.ts
import { viteStaticCopy } from "vite-plugin-static-copy";

export default {
  plugins: [
    viteStaticCopy({
      targets: [
        {
          src: "node_modules/@marcusok/pdf-preview/dist/cmaps/*",
          dest: "pdf/cmaps",
        },
        {
          src: "node_modules/@marcusok/pdf-preview/dist/standard_fonts/*",
          dest: "pdf/standard_fonts",
        },
      ],
    }),
  ],
};
```

```ts
createPdfPreview(el, {
  source,
  cMapUrl: "/pdf/cmaps/",
  standardFontDataUrl: "/pdf/standard_fonts/",
});
```

If the app is deployed under a sub-path, prefix these URLs with your base (`import.meta.env.BASE_URL`, or VitePress's `withBase`): a root-absolute `/pdf/cmaps/` is only correct at the domain root.

Missing cmaps do **not** throw — CJK documents silently render blank/garbled text, which is why this wiring matters. A quick smoke test with a real Chinese PDF is the only reliable verification.

## Scaling

```ts
createPdfPreview(el, { source, scale: "fit-width" }); // default
createPdfPreview(el, { source, scale: "fit-page" });
createPdfPreview(el, { source, scale: 1.5 }); // raw viewport scale
```

`fit-*` modes measure the container once at creation. A hidden container (zero width) falls back to `scale = 1` with a `console.warn`. Container resizes do not re-fit — destroy and re-create the instance.

## Next steps

- [Limits](./02-limits) — memory behavior, wasm-gated features, CSP requirements, SSR boundaries.
- Try it with your own files in the [play](/play).
