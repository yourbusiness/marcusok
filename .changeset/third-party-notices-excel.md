---
"@marcusok/excel-exporter": patch
"@marcusok/excel-preview": patch
---

Ship `THIRD_PARTY_NOTICES.md` with the bundled third-party software (modern-xlsx — engine JS glue plus the re-shipped WASM binary — and fflate, both MIT), matching the word/pdf/ppt preview packages. MIT requires the copyright and permission notice to accompany substantial portions of the software, and the packages' `dist` embeds both.
