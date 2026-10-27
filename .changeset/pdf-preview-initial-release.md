---
"@marcusok/pdf-preview": minor
---

Initial release: read-only PDF preview in the browser, powered by pdf.js (pdfjs-dist 6.4.299 bundled — zero transitive dependencies).

- `createPdfPreview(container, options)` one-line mounting with continuous scroll and viewport-driven lazy rendering (pages render as they approach the viewport, canvases release beyond 3 screens, renders serialized per instance; per-page placeholders sized from each page's own viewport so mixed-size documents don't jump)
- Selectable/copyable text layer (pdf.js `TextLayer`) with the required positioning CSS injected as an idempotent singleton (`styleNonce` supported); `textLayer: false` for pure canvas
- `scale`: number / `"fit-width"` (default) / `"fit-page"`, resolved once at creation
- Worker + cmaps + standard_fonts ship in `dist`; under Vite/webpack the worker needs zero config (bundler-emitted anchor), cmaps/fonts are URL-overridable for bundler scenarios (CJK documents degrade silently without them)
- Format sniffing before the engine (`%PDF-` magic within the first 1024 bytes, Adobe's reader window) and normalized error codes: `UNSUPPORTED` / `PASSWORD_PROTECTED` (detected, not decrypted) / `CORRUPT` / `UNKNOWN`
- Hardened defaults: the pdf.js sandbox module is never loaded (embedded PDF JavaScript never executes), the annotation layer is not rendered, XFA is off
- `destroy()` full-chain cancellation: observers disconnected, in-flight render tasks and text-layer streams cancelled, rendered canvases zeroed, worker terminated, DOM detached — destroy-path rejections are silent
- Worker / asset load failures report `UNKNOWN` (with a worker-specific message) instead of masquerading as a corrupt PDF; explicit `workerSrc` may be absolute or relative (resolved against the page base before deriving cmaps/fonts)
