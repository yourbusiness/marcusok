---
"@marcusok/word-preview": patch
---

Three fixes: (1) the normalized "render failed" message for engine null-reference crashes now phrases the missing-relationship-part cause as _a common cause_ instead of asserting it — other null crashes inside the engine produced the same message and were misattributed to a broken file. (2) Anchors with leading C0/space characters (`href=" #section"`) are preserved on non-http(s) pages (file://, custom protocols), matching how browsers strip those characters before resolving the href. (3) The package re-enables the repo-wide `noUncheckedIndexedAccess` TypeScript check (the ZIP flag sniff in production code now reads the flag byte defensively).
