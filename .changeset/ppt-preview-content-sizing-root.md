---
"@marcusok/ppt-preview": patch
---

Scroll-root detection now also rejects content-sizing keywords (`max-content`, `min-content`, `fit-content` including parameterized and legacy-prefixed variants, `stretch`, `-moz-available`, `-webkit-fill-available`) for `height`/`max-height`. A wrapper sized by these keywords grows with its content, so selecting it as the IntersectionObserver root made the render/release bands cover the whole document — lazy rendering degraded to full rendering and `goToSlide`/`nextSlide`/`prevSlide` scrolled a container that never scrolls (silent no-op). Rejecting a genuinely definite root only falls back to the viewport root where lazy rendering and navigation keep working, so the check is deliberately conservative.
