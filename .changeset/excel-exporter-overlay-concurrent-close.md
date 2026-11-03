---
"@marcusok/excel-exporter": patch
---

Re-render the progress overlay immediately when a concurrent caller closes early while the overlay is already revealed (bundled `progress-overlay` fix). The retired caller's title/label previously stayed on screen for the remainder of the still-active task whenever that task emitted no further progress/phase events — the symmetric gap of the show-takeover re-render fix in 3.4.5.
