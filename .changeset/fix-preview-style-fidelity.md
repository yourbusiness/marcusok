---
"@marcusok/excel-preview": patch
---

Fix a set of style-restoration and robustness issues found in a full review:

- Underlines are no longer lost on fonts written as `<u val="single"/>` (the explicit default form, common in third-party exports such as LibreOffice) — the val attribute is an enum, not a boolean.
- Diagonal borders now respect the `diagonalUp`/`diagonalDown` flags: direction is rendered accordingly (`"/"` vs `"\"`), and a `<diagonal style>` without either flag no longer draws a line Excel itself would not show.
- Gradient fills and diagonal borders no longer clobber each other — both land in one layered `background-image` declaration (diagonal on top).
- Bold/italic are preserved when underline and strike are combined into one `text-decoration`.
- A cross-origin `parseWorkerUrl` (raw CDN URL) now falls back to main-thread parsing with a console warning instead of throwing `SecurityError` out of the parse call; the docs now state that worker scripts are same-origin only.
- A hung parse worker (WASM init stuck / engine loop) is terminated after `workerTimeoutMs` (default 120s), its in-flight requests are rejected without a main-thread rerun, and the next request rebuilds a fresh worker.
