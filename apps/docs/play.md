# Play

Everything on this page runs in your browser against the packages as published — parsing, building and progress are all local, nothing is uploaded. Data comes from the docs site's deterministic mock generator (mulberry32 seeded PRNG), so repeated runs with the same settings produce identical output.

## Excel export — `@marcusok/excel-exporter`

Pick a dataset, row count and export mode, click export to get a real `.xlsx` file, and watch progress, phase timings and the actual engine used.

<ClientOnly>
  <PackageDemo dir="excel-exporter" />
</ClientOnly>

- **auto (recommended)**: picks the optimal main / worker / stream path by row and cell count;
- **main**: synchronous main-thread build — notice the cliff at 100k rows;
- **worker**: Web Worker threading; the main thread only does one structured clone;
- **stream**: Fast stream, ~0.8s at 100k rows, but no styles or layout features.

Full usage docs: [excel-exporter](/packages/excel-exporter/).

## xlsx preview — `@marcusok/excel-preview`

Drop an `.xlsx` / `.xlsm` / `.csv` file on the panel: the parse runs in a Web Worker and the grid renders with DOM virtual scrolling, so the page stays responsive on large workbooks. Per-stage timings (parse / render / total) are reported as soon as the file resolves; an encrypted workbook reports `PASSWORD_PROTECTED` instead of failing silently.

<ClientOnly>
  <PackageDemo dir="excel-preview" />
</ClientOnly>

Full usage docs: [excel-preview](/packages/excel-preview/).

## Word preview — `@marcusok/word-preview`

Drop a `.docx` / `.docm` file on the panel: the document renders with visual fidelity (pagination, headers/footers, footnotes as saved by Word), with per-render timing reported via `onRendered`. Rendering runs on the main thread (an engine constraint — there is no worker mode); untrusted-input hardening is on by default (altChunks off, hyperlink scheme allowlist). A synthetic 1000-paragraph document and a two-instance isolation demo are built into the panel.

<ClientOnly>
  <PackageDemo dir="word-preview" />
</ClientOnly>

Full usage docs: [word-preview](/packages/word-preview/).

## PDF preview — `@marcusok/pdf-preview`

Drop a `.pdf` file on the panel: parsing runs in a pdf.js Web Worker and pages render lazily as they scroll into view (far-away pages release their canvas memory), with first-paint and end-to-end timings reported via `onRendered`. The transparent text layer makes the preview selectable and copyable; encrypted PDFs report `PASSWORD_PROTECTED`. On this site the CMap / standard-font directories are served from `public/` (the bundler only emits the worker).

<ClientOnly>
  <PackageDemo dir="pdf-preview" />
</ClientOnly>

Full usage docs: [pdf-preview](/packages/pdf-preview/).

## PPT preview — `@marcusok/ppt-preview`

Drop a `.pptx` file on the panel: parsing runs in a Web Worker (Rust/WASM engine) and slides render lazily as they approach the viewport (far-away slides release their canvas memory), with first-slide and end-to-end timings reported via `onRendered`. The transparent text layer makes slides selectable and copyable; navigation is 0-based (`goToSlide` / `nextSlide` / `prevSlide`), and encrypted files (MS-OFFCRYPTO Agile) decrypt when a password is provided.

<ClientOnly>
  <PackageDemo dir="ppt-preview" />
</ClientOnly>

Full usage docs: [ppt-preview](/packages/ppt-preview/).

## Progress overlay — built into `@marcusok/excel-exporter`

The overlay the exporter shows during an export, driven here directly (through the `/overlay` subpath) for a task that has nothing to do with Excel. The simulated task reports no progress for its first 1.5 seconds — that is the indeterminate spinner — then streams progress and switches to the percentage bar. It is the same two-state panel, on the same implementation the `overlay` option uses.

<ClientOnly>
  <OverlayDemo />
</ClientOnly>

Full usage docs: [Progress overlay guide](/packages/excel-exporter/guide/11-overlay).
