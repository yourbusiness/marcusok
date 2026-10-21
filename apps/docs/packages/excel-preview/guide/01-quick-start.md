# Quick Start

One `createPreview` call mounts a full preview: worker parsing, error normalization and the virtual-scrolling renderer.

## Minimal example

```ts
import { createPreview } from "@marcusok/excel-preview";

const el = document.querySelector("#preview")!;
const preview = createPreview(el, {
  source: file, // File | Blob | Uint8Array | ArrayBuffer
});

// later
preview.destroy();
```

That's the entire setup. Parsing happens in a Web Worker by default; the same pipeline runs on the main thread — no code change — whenever no worker is usable: Node/SSR (no Worker **or** no `window` global), or a worker that failed to construct/load (a cross-origin script URL, a 404 — the fallback is then permanent for the session; see [Assets & Self-hosting](/packages/excel-preview/guide/02-assets)).

Give the container a definite height (e.g. `#preview { height: 500px; }` or a flex/grid track). The preview root is `height: 100%` with a 200px minimum and scrolls inside the container; a container without a resolved height falls back to that minimum.

## Options you will actually use

```ts
const preview = createPreview(el, {
  source: bytes,
  password: "…", // encrypted workbooks (Agile AES-256)
  sheet: "Summary", // initial sheet by name or 0-based index
  showHeaders: false, // hide the A/B/C + 1/2/3 headers
  showGridLines: false, // override the file's gridline switch
  showTabs: false, // hide the sheet tab bar
  onParsed: (info) => {
    // sheetNames / sheetCount / rowCount / colCount / duration{parse,render,total}
  },
  onError: (e) => {
    // e.code: PASSWORD_PROTECTED | LEGACY_FORMAT | CORRUPT | UNSUPPORTED | WASM | UNKNOWN
  },
});
```

## Switching sheets

```ts
preview.setSheet(1); // by index
preview.setSheet("Sheet2"); // by name
preview.getSheetNames(); // all sheets, file order (hidden ones included)
```

The tab bar (rendered by default) does the same thing visually; sheets marked `hidden` / `veryHidden` in the file never appear as tabs.

Timing matters: these methods only take effect **after the parse resolves**. `createPreview` boots asynchronously, so calling `setSheet()` on the line right after it returns is a silent no-op — no return value, no `onError` — and `getSheetNames()` returns `[]` until then. Call them from `onParsed`, from your own UI, or let the tab bar drive them. After boot, invalid values (unknown name, out-of-range index, non-integer number, `NaN`) report `Sheet not found: …` through `onError` instead of throwing. After `destroy()` both are no-ops again; note that `destroy()` **also aborts an in-flight parse** — the boot path returns silently as `ABORTED` (`onError` / `onParsed` never fire for a cancelled preview; see the API reference for the cancellation semantics).

The container may live inside a ShadowRoot **or an iframe document**: styles are injected into the container's root node / `ownerDocument`, and an optional `styleNonce` (renderer option / `createPreview` option) is applied to the injected `<style>` elements for CSP setups.

## The low-level parse API

When you don't want the renderer at all — custom UI, SSR, or a framework wrapper:

```ts
import { parseWorkbookBytes } from "@marcusok/excel-preview";

const workbook = await parseWorkbookBytes(bytes, { password: "…" });
// workbook.sheets[0].rows[0].cells[0] → { col, type, value, styleIndex }
// styles: fonts / fills / borders / xfs (numFmtCode already resolved)
```

The model is plain JSON (structured-clone safe) and is exactly what the worker posts back to the renderer. See the [model types](/packages/excel-preview/api/02-model).

## Next steps

- [Assets & self-hosting](/packages/excel-preview/guide/02-assets) — how the WASM/worker files resolve, and when you need `configureWasm`.
- [Format fidelity](/packages/excel-preview/guide/03-format-fidelity) — what the compensation layer restores and the known approximations.
- [Limits](/packages/excel-preview/guide/04-limits) — what the package deliberately does not do.
