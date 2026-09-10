# Getting Started

Run your first Excel export in minutes. Requirement: Node `>= 22`. Example commands use pnpm; npm / yarn work the same (`pnpm >= 9` is only a dev requirement of this repository, not of consumers).

## 1. Install

```bash
pnpm add @marcusok/excel-exporter
```

That is the entire setup — one package, zero runtime dependencies. The export engine (modern-xlsx JS glue + fflate) is bundled in at build time, and the WASM binary ships under this package's own `exports` map, so there is no engine package to install and no `engines` conflict from upstream ranges.

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
        { key: "orderId", header: "Order ID", width: 18 },
        { key: "date", header: "Date", width: 12, format: { type: "date" } },
        {
          key: "amount",
          header: "Amount",
          width: 14,
          style: StylePresets.currency,
        },
        {
          key: "status",
          header: "Status",
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

No `main.ts` wiring, no bundler plugins: the two shipped assets (`modern-xlsx.wasm`, `export.worker.js`) are located automatically — bundlers emit them as hashed assets via the standard `new URL(asset, import.meta.url)` pattern, and Node reads the wasm from disk. Node / SSR environments need no browser assets and no initialization boilerplate (see [Node/SSR](/packages/excel-exporter/guide/09-node-ssr)). Self-hosted or CDN-hosted copies are the one case that needs [`configureWasm`](/packages/excel-exporter/guide/02-installation).

## 3. Next steps

- Learn how [auto mode routing](/packages/excel-exporter/guide/03-auto-mode) picks main / worker / stream
- Try different modes and row counts in the [play](/play)
