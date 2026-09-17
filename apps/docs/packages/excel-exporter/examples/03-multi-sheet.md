# Example: Multi-Sheet Workbook

Admin reports often combine several topics in one file. This example builds a workbook from two mock datasets: sales detail and a staff roster.

## Mock data preview (staff roster)

<MockPreview dataset="staff" :rows="5" />

## Implementation

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

// salesRows / staffRows: data from your business layer (fetching is omitted
// here); fields match each sheet's columns below

const result = await exportExcel({
  filename: "department-report-2026-Q3",
  sheets: [
    {
      name: "Sales",
      freezeRows: 1,
      autoFilter: true,
      columns: [
        { prop: "orderId", label: "Order ID", width: 18 },
        { prop: "product", label: "Product", width: 18 },
        {
          prop: "amount",
          label: "Amount",
          width: 14,
          style: StylePresets.currency,
        },
      ],
      data: salesRows,
    },
    {
      name: "Staff",
      freezeRows: 1,
      columns: [
        { prop: "id", label: "ID", width: 10 },
        { prop: "name", label: "Name", width: 12 },
        { prop: "dept", label: "Department", width: 12 },
        { prop: "position", label: "Position", width: 14 },
        {
          prop: "salary",
          label: "Monthly Salary",
          width: 14,
          style: StylePresets.currency,
        },
        {
          prop: "hiredAt",
          label: "Hired",
          width: 12,
          format: { type: "date" },
        },
      ],
      data: staffRows,
    },
  ],
});
```

## Notes

- Each sheet defines its own columns, styles and data independently;
- Total rows = the sum across sheets; `auto` routes on the **total**;
- Sheet names follow Excel rules: ≤ 31 characters, no `: \ / ? * [ ]`.
