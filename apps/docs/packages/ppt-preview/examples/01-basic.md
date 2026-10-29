# Basic Example

A minimal, complete wiring: pick a file, preview it, navigate, clean up.

```ts
import { createPptPreview } from "@marcusok/ppt-preview";
import type {
  PptPreviewInstance,
  PptPreviewError,
} from "@marcusok/ppt-preview";

const container = document.querySelector<HTMLDivElement>("#preview")!;
const fileInput = document.querySelector<HTMLInputElement>("#file")!;
const status = document.querySelector<HTMLElement>("#status")!;

let preview: PptPreviewInstance | null = null;

fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (!file) return;

  // Destroy the previous instance before mounting a new one.
  preview?.destroy();
  preview = null;
  container.textContent = "";

  preview = createPptPreview(container, {
    source: file,
    slideWidth: 960,
    onRendered: ({ slideCount, firstSlideMs, durationMs }) => {
      status.textContent = `${slideCount} slides · first ${firstSlideMs}ms · total ${durationMs}ms`;
    },
    onError: (e: PptPreviewError) => {
      status.textContent = `${e.code}: ${e.message}`;
    },
  });
});

// 0-based navigation wired to buttons:
document
  .querySelector("#prev")
  ?.addEventListener("click", () => preview?.prevSlide());
document
  .querySelector("#next")
  ?.addEventListener("click", () => preview?.nextSlide());
document
  .querySelector("#first")
  ?.addEventListener("click", () => preview?.goToSlide(0));
```

Markup — the container must actually scroll (definite height + `overflow: auto`):

```html
<input id="file" type="file" accept=".pptx" />
<button id="first">First</button>
<button id="prev">Prev</button>
<button id="next">Next</button>
<span id="status"></span>
<div id="preview" style="height: 600px; overflow: auto"></div>
```

## Encrypted file variant

```ts
preview = createPptPreview(container, {
  source: encryptedFile,
  password: passwordInput.value, // MS-OFFCRYPTO Agile "password to open"
  onRendered,
  onError: (e) => {
    // e.code === "PASSWORD_PROTECTED"; e.message distinguishes
    // "supply a password" from "the password was rejected".
  },
});
```

## Framework teardown

Destroy on unmount (React example):

```tsx
useEffect(() => {
  if (!file || !ref.current) return;
  const p = createPptPreview(ref.current, { source: file });
  return () => p.destroy();
}, [file]);
```
