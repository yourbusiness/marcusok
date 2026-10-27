# Basic Usage

## Upload → preview (React)

```tsx
import { useEffect, useRef, useState } from "react";
import {
  createPdfPreview,
  type PdfPreviewInstance,
} from "@marcusok/pdf-preview";

export function PdfPanel({ file }: { file: File }) {
  const ref = useRef<HTMLDivElement>(null);
  const previewRef = useRef<PdfPreviewInstance | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!ref.current) return;
    const preview = createPdfPreview(ref.current, {
      source: file,
      scale: "fit-width",
      onError: (e) => setError(`${e.code}: ${e.message}`),
    });
    previewRef.current = preview;
    return () => preview.destroy(); // cancels in-flight work + worker
  }, [file]);

  return (
    <div style={{ height: 600, overflow: "auto", background: "#f5f5f5" }}>
      {error && <p>{error}</p>}
      <div ref={ref} />
    </div>
  );
}
```

The scroll container is your own element — the preview grows inside it and the viewport-driven lazy rendering follows along automatically. The observers root at the nearest ancestor that actually scrolls or has a constrained height (fixed height, `max-height`, or a flex-resolved height), falling back to the browser viewport; an `overflow: auto` wrapper that merely grows with its content is skipped, so the preload/release bands track the element that really scrolls.

## Scale modes

```ts
createPdfPreview(el, { source, scale: "fit-width" }); // page width = container width (default)
createPdfPreview(el, { source, scale: "fit-page" }); // whole page fits the container
createPdfPreview(el, { source, scale: 1.25 }); // raw pdf.js viewport scale
```

`fit-*` resolves once at creation. To change the scale later, `destroy()` and create a new instance — v1 intentionally ships no hot re-layout API.

## Text layer on/off

```ts
createPdfPreview(el, { source, textLayer: false }); // pure canvas (no selection)
```

The text layer is on by default: transparent spans exactly aligned with the canvas glyphs, so selection and copy work like real text.

## Performance instrumentation

```ts
createPdfPreview(el, {
  source,
  onRendered: ({ pageCount, firstPageMs, durationMs }) => {
    // pageCount: document total; firstPageMs: first painted page (canvas + text layer);
    // durationMs: from the createPdfPreview call to that first paint
  },
});
```

For a 200-page synthetic text document, parse + first-page text extraction measures ~66 ms locally (Node baseline; see the package README) — the number end users care about is `firstPageMs`, which includes canvas paint.
