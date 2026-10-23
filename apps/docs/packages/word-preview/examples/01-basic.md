# Basic Usage

## Render an uploaded file

```ts
import { createWordPreview } from "@marcusok/word-preview";

const input = document.querySelector<HTMLInputElement>("#file")!;

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;

  const preview = createWordPreview(document.querySelector("#preview")!, {
    source: file,
    onRendered: ({ pageCount, durationMs }) =>
      console.log(`${pageCount} sections, full render ${durationMs}ms`),
    onError: (e) => {
      switch (e.code) {
        case "LEGACY_FORMAT":
          showHint(
            "Legacy .doc and encrypted .docx are both unsupported — re-save as .docx",
          );
          break;
        case "UNSUPPORTED":
          showHint("This is not a Word document");
          break;
        default:
          showHint(e.message);
      }
    },
  });

  // later: preview.destroy()
});
```

## Render from bytes with render options

```ts
import { createWordPreview } from "@marcusok/word-preview";

const bytes: Uint8Array = await response
  .arrayBuffer()
  .then((b) => new Uint8Array(b));

const preview = createWordPreview(el, {
  source: bytes,
  renderOptions: {
    // whitelisted engine options — 13 of them, see the API page
    ignoreWidth: true,
    renderComments: false,
    // flip the package default back to the engine's own behavior
    ignoreLastRenderedPageBreak: true,
  },
});
```

## Two documents on one page

Instances isolate their styles automatically — just create two:

```ts
const a = createWordPreview(leftPane, { source: fileA });
const b = createWordPreview(rightPane, { source: fileB });
// each renders under its own docx-<n> class; numbering/theme CSS never crosses
```

## React-style teardown

`destroy()` is synchronous and idempotent; a render still in flight writes into a detached node and is dropped:

```tsx
useEffect(() => {
  const preview = createWordPreview(ref.current, { source });
  return () => preview.destroy();
}, [source]);
```
