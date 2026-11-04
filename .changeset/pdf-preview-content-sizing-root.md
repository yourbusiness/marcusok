---
"@marcusok/pdf-preview": patch
---

Scroll-root detection catches two missed non-definite-height forms: `calc()` max-heights containing a percentage (e.g. `calc(100% - 2rem)`, previously only values _ending_ with `%` were rejected) and content-sizing keywords (`max-content`, `min-content`, `fit-content` variants, `stretch`). Either form as a wrapper's `height`/`max-height` made the IO root span the whole document — lazy rendering degraded to full rendering and canvases were never released. The package also re-enables the repo-wide `noUncheckedIndexedAccess` TypeScript check (production code already passed; test files were updated).
