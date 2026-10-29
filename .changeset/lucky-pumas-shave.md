---
"@marcusok/ppt-preview": minor
---

Add `@marcusok/ppt-preview` — read-only .pptx preview in the browser. Continuous scroll with lazy per-slide Canvas rendering (viewport-adjacent slides render, far-away slides release canvas memory), a selectable/copyable text layer collected in the same render pass, 0-based slide navigation (`goToSlide`/`nextSlide`/`prevSlide`) tracking the most visible slide, password decryption for MS-OFFCRYPTO Agile encrypted files, and normalized error codes. Powered by a Rust/WASM engine (@silurus/ooxml, bundled at build time — zero transitive dependencies; the parser wasm ships in `dist`).
