---
"@marcusok/excel-preview": patch
---

Make the published `.d.ts` fully self-contained by inlining the two engine type references (`BorderStyle`, `AlignmentData`, now also exported for custom renderers), and move `modern-xlsx` back to devDependencies. Consumers no longer install `modern-xlsx` as a transitive runtime dependency, which also drops its `engines.node >= 24` declaration from consumer installs (the runtime engine was already bundled into dist; nothing changes at runtime).
