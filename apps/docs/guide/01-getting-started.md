# Getting Started

MarcusOK publishes four packages across three categories — how they relate is on the [Ecosystem](/guide/) page. This page starts with the one decision that comes first (which package to install), then walks through a first export with the most common one.

Requirement: Node `>= 22`. Example commands use pnpm; npm / yarn work the same (`pnpm >= 9` is only a dev requirement of this repository, not of consumers).

## Which package do I need?

| I want to…                                                     | Install                      | Start at                                                            |
| -------------------------------------------------------------- | ---------------------------- | ------------------------------------------------------------------- |
| Export data to a downloadable `.xlsx`                          | `@marcusok/excel-exporter`   | The [exporter walkthrough](#_2-first-export) below                  |
| Show an uploaded `.xlsx` / `.xlsm` / `.csv`, read-only         | `@marcusok/excel-preview`    | [Preview quick start](/packages/excel-preview/guide/01-quick-start) |
| Drive the same progress overlay for a non-Excel task           | `@marcusok/progress-overlay` | [Overlay usage](/packages/progress-overlay/guide/01-usage)          |
| Build directly on the modern-xlsx engine (own reader / writer) | `@marcusok/xlsx-core`        | [Loader guide](/packages/xlsx-core/guide/01-loader)                 |

Everything below walks through `@marcusok/excel-exporter`; the other packages' quick starts are linked in the table above.

Two of those rows are unusual in practice: `@marcusok/xlsx-core` and `@marcusok/progress-overlay` are normally installed **for** you as dependencies of the two document packages, so you only add them explicitly when you want them on their own. The full comparison — dependency graph, what is shared at runtime, how versions relate — is in [Package Relationships & Selection](/guide/03-package-relationships).

## 1. Install

```bash
pnpm add @marcusok/excel-exporter
```

That is the entire setup — two same-scope runtime dependencies (`@marcusok/xlsx-core`, the shared engine layer; `@marcusok/progress-overlay`, the shared overlay UI since 2.8.0). The export engine (modern-xlsx JS glue + fflate) is bundled in at build time inside the core layer, and the WASM binary ships under the core package's `exports` map, so there is no engine package to wire in and no `engines` conflict from upstream ranges.

## 2. First export

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "sales-report-2026",
  sheets: [
    {
      name: "Sales",
      freezeRows: 1,
      autoFilter: true,
      columns: [
        { prop: "orderId", label: "Order ID", width: 18 },
        { prop: "date", label: "Date", width: 12, format: { type: "date" } },
        {
          prop: "amount",
          label: "Amount",
          width: 14,
          style: StylePresets.currency,
        },
        {
          prop: "status",
          label: "Status",
          width: 10,
          format: {
            type: "enum",
            map: { paid: "Paid", pending: "Pending" },
            fallback: "Unknown",
          },
        },
      ],
      data: [
        {
          orderId: "ORD-000001",
          date: "2026-07-01",
          amount: 1299.99,
          status: "paid",
        },
      ],
    },
  ],
});
```

In the browser this triggers a download; `.xlsx` is appended when missing. Use `download: false` to receive the Blob only.

No `main.ts` wiring, no bundler plugins: the shipped assets (`modern-xlsx.wasm` via the shared `@marcusok/xlsx-core`, plus this package's `export.worker.js`) are located automatically — bundlers emit them as hashed assets via the standard `new URL(asset, import.meta.url)` pattern, and Node reads the wasm from disk. Node / SSR environments need no browser assets and no initialization boilerplate (see [Node/SSR](/packages/excel-exporter/guide/09-node-ssr)). Self-hosted or CDN-hosted copies are the one case that needs [`configureWasm`](/packages/excel-exporter/guide/02-installation).

## 3. Next steps

- Learn how [auto mode routing](/packages/excel-exporter/guide/03-auto-mode) picks main / worker / stream
- Try different modes and row counts in the [play](/play) — the same page also demos the preview package and the progress overlay
- Start on another package from the [table above](#which-package-do-i-need), or see [Package Relationships & Selection](/guide/03-package-relationships) to compare them
