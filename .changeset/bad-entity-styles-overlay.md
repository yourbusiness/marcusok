---
"@marcusok/excel-preview": patch
---

Harden the styles-overlay XML entity decoder against malicious character references: an out-of-range numeric entity (`&#1114112;`) previously threw a `RangeError` that escaped the overlay's degradation guard and failed the whole parse with an `UNKNOWN` error, instead of falling back to the engine style arrays. Such entities (and lone surrogates) are now replaced with U+FFFD, and the hexadecimal character-reference form (`&#x41;`) is decoded per the XML spec.
