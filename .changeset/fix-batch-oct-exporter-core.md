---
"@marcusok/excel-exporter": patch
---

Hardened input validation, cross-path cell normalization, and worker payload fixes:

- Style colors (`font.color`, `fill.fgColor`/`bgColor`, border side colors) are validated as 6-digit RGB hex and border styles against the OOXML enum — invalid values now fail fast with a structured `{ success: false }` instead of `success: true` with a workbook Excel may flag as corrupt.
- A function-form format returning `undefined`/`null` now yields a physical blank cell on the workbook path, matching the stream path (`ISBLANK()`/`COUNTA()` no longer drift when the same data crosses the 50k-row routing threshold or degrades).
- The `overlay` option is stripped from the worker postMessage payload: it is consumed entirely on the main thread and may legally carry DOM values (e.g. `container: HTMLElement`), which made every worker-route export throw `DataCloneError` and rebuild on the main thread.
- The bundled engine layer (xlsx-core) now logs the Node WASM auto-init root cause (ENOENT / CompileError) before degrading, instead of swallowing it and later failing with a misleading "fetch failed".
