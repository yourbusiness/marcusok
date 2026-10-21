# Example: 100k-Row Export

Large exports are what the Fast stream path is built for: 100k rows run in a Worker in the browser (the main thread only does one structured clone), with Fast stream finishing in ~0.8s.

## Mock data preview

<MockPreview dataset="sales" :rows="3" />

## Implementation

```ts
import { exportExcel } from "@marcusok/excel-exporter";

// rows: sales data from your business layer, fields matching the columns
// below (fetching is omitted here; the scenario uses 100k rows — the live
// demo generates mock data for the same sales scenario, with more columns)

const result = await exportExcel({
  filename: "large-export-100k",
  sheets: [
    {
      name: "Sales",
      columns: [
        { prop: "orderId", label: "Order ID", width: 18 },
        { prop: "date", label: "Date", width: 12 },
        { prop: "amount", label: "Amount", width: 14 },
        { prop: "status", label: "Status", width: 10 },
      ],
      data: rows,
    },
  ],
  mode: "auto", // at/above the row/cell gate -> worker + stream
  onProgress: (p) => setProgress(p),
});

console.log(result); // engine: "fast-xlsx", mode: "stream", rowCount: 100000
```

## Notes

- At 100k rows `auto` picks `worker + Fast stream`: ~0.8s measured (vs 17.5s on the Workbook path in Chrome);
- The stream path **supports multi-row headers, merges and number formats** (date/number FormatSpec via minimal `styles.xml`), but does **not** support cell/layout styles, width/freeze/filter; a console warning is expected;
- `onProgress` reports every 1000 rows;
- `decimals`/`thousands` are display-only numFmt settings on both paths — stored values stay identical across the threshold.

Compare `auto` vs `main` at 100,000 rows in the [play](/play).
