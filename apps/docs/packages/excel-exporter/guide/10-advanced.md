# Advanced Features

## Multiple sheets

`sheets` is an array — one call produces a multi-page workbook:

```ts
await exportExcel({
  filename: "department-report",
  sheets: [
    { name: "Sales", columns: [...], data: salesData },
    { name: "Staff", columns: [...], data: staffData },
  ],
});
```

Sheet names must satisfy ECMA-376 constraints: non-empty, ≤ 31 characters, must not contain `: \ / ? * [ ]`, and must not begin or end with an apostrophe. They must also be unique across `sheets` — a duplicate name fails the same way instead of producing a corrupt file or a silently renamed sheet. A violation never produces a corrupt file and never throws to the caller: input is validated up front, before any export route (or the fallback chain) runs, and the call resolves with `{ success: false, error }` carrying a clear message.

## Frozen rows

`freezeRows: 1` freezes the header row (mapped to `frozenPane`). For multi-row headers, prefer `freezeRows >= header row count` so all header rows stay visible.

## Multi-row headers

Nest columns with `children` to get multi-row headers. Group header cells merge across all their descendant leaf columns, and leaf headers span the remaining header rows vertically — no manual merge math required.

```ts
await exportExcel({
  filename: "monthly-sales",
  sheets: [
    {
      name: "Sales",
      freezeRows: 3,
      columns: [
        { key: "product", header: "Product" },
        {
          header: "Revenue",
          children: [
            {
              header: "This month",
              children: [
                { key: "m_qty", header: "Qty" },
                { key: "m_amt", header: "Amount" },
              ],
            },
            {
              header: "YTD",
              children: [
                { key: "y_qty", header: "Qty" },
                { key: "y_amt", header: "Amount" },
              ],
            },
          ],
        },
      ],
      data: [{ product: "A", m_qty: 1, m_amt: 2, y_qty: 3, y_amt: 4 }],
    },
  ],
});
```

This produces a 3-row header:

| Row | A       | B                       | C      | D                | E      |
| --- | ------- | ----------------------- | ------ | ---------------- | ------ |
| 1   | Product | Revenue (merged B–E)    |        |                  |        |
| 2   |         | This month (merged B–C) |        | YTD (merged D–E) |        |
| 3   |         | Qty                     | Amount | Qty              | Amount |

Rules:

- Leaf columns (no `children`) need a `key`; group columns may omit it and contribute header rows only;
- `width` / `style` / `format` apply to leaf columns only;
- Group header cells style via that column's `headerStyle`, leaf headers likewise (falling back to the sheet-level `headerStyle`);
- Multi-row headers work on every path (main / worker / stream, including the stream fallback); merges survive on the stream and fallback paths too (styles excepted).

Try it live: below is a mock preview of the `sales-grouped` dataset (two-level grouped header, matching the exported file's header structure); pick the `sales-grouped` dataset in the [playground](/play) demo panel to export a real file with a multi-row header and data-area merges.

<MockPreview dataset="sales-grouped" :rows="5" />

## Merged cells

```ts
{
  name: "Summary",
  columns: [...],
  data: [...],
  merges: [
    { row: 0, col: 0, rowspan: 1, colspan: 2 }, // first data row spans two columns
  ],
}
```

`MergeRange` is relative to the data area: `row` / `col` start at 0 (`row 0` = first data row); `rowspan` / `colspan` are the spans.

Merges are validated identically on every path: values must be integers, `row`/`col` ≥ 0, `rowspan`/`colspan` ≥ 1, the range must stay inside the data area (leaf column count / data row count), and ranges must not overlap each other. Invalid input resolves with `{ success: false, error }` naming the offending merge — never a workbook Excel flags as corrupt.

## Auto filter

`autoFilter: true` adds filter dropdowns to the header range.

## Progress and phase callbacks

```ts
await exportExcel({
  ...,
  onProgress: (progress) => {
    // 0 → 1; the leading 0 and trailing 1 fire exactly once each on every route
    // (the stream fallback included); incremental progress only on the stream
    // path (every 1000 rows)
    bar.style.width = `${progress * 100}%`;
  },
  onPhase: (phase, durationMs) => {
    // phase: "init" | "build" | "download", strictly sequential
    console.log(`${phase} took ${durationMs.toFixed(1)}ms`);
  },
});
```

Phase semantics:

| Phase      | Description                                                                                                                                                                                                                                                                                                   |
| ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `init`     | WASM init; reported on every main-path export (~0ms once loaded); Node's main-thread stream path and the main-thread stream fallback do not load WASM but still report a single 0ms to keep the phase sequence stable; on Worker + Workbook only when the worker initializes; not reported on Worker + stream |
| `build`    | Workbook construction (reported once per actual attempt, including fallback)                                                                                                                                                                                                                                  |
| `download` | Browser download trigger (absent with `download: false`; absent in Node)                                                                                                                                                                                                                                      |

> `onPhase` measures per-phase wall time only; `ExportResult.duration` measures the whole export on main-thread routes (the worker route's duration covers the in-worker time only).

## Disable auto download

```ts
const result = await exportExcel({ ..., download: false });
// use result.blob directly
```

## Export result

```ts
interface ExportResult {
  success: boolean;
  blob?: Blob;
  engine?: "modern-xlsx"; // engine actually used
  mode?: ExportMode; // mode actually used
  duration?: number; // total export duration in ms
  rowCount?: number;
  error?: Error;
}
```

Show `result.error` on failure; on a successful export a non-empty `result.error` marks a degraded (style-less stream) export.
