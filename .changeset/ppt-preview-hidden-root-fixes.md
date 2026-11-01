---
"@marcusok/ppt-preview": patch
---

Re-detect the scroll root once a container that was hidden at creation becomes visible, so `goToSlide`/`nextSlide`/`prevSlide` scroll the real panel instead of the window and lazy preloading works; also correct the password-error, `dpr` clamp, Vite production-build, hidden-container `onRendered`, and resource-budget wording in the documentation.
