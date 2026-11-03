---
"@marcusok/ppt-preview": patch
---

Exclude `calc(...)`-wrapped percentage lengths (`calc(100% - 2rem)`) from the scroll-root definite-height heuristic. They previously passed the `endsWith("%")` check, were treated as a definite height, and could be picked as the IntersectionObserver root — reproducing the 0.1.6-class symptom (root spanning the whole deck, full rendering, navigation silently no-op). Plain percentage and `vh`-based values keep their previous behavior; add regression coverage for the `wasmUrl` / `useGoogleFonts` load-options wiring.
