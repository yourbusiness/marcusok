# 案例：多 Sheet 工作簿

后台报表经常需要一份文件包含多个主题的 Sheet。本案例用销售明细 + 人员花名册两个 mock 数据集生成多页工作簿。

## Mock 数据预览（人员花名册）

<MockPreview dataset="staff" :rows="5" />

## 实现代码

```ts
import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

// salesRows / staffRows：业务侧数据（由你的业务代码提供，此处省略取数过程），
// 字段分别与两个 sheet 的 columns 一一对应

const result = await exportExcel({
  filename: "部门经营快报-2026-Q3",
  sheets: [
    {
      name: "销售明细",
      freezeRows: 1,
      autoFilter: true,
      columns: [
        { prop: "orderId", label: "订单号", width: 18 },
        { prop: "product", label: "商品", width: 18 },
        {
          prop: "amount",
          label: "金额",
          width: 14,
          style: StylePresets.currency,
        },
      ],
      data: salesRows,
    },
    {
      name: "人员花名册",
      freezeRows: 1,
      columns: [
        { prop: "id", label: "工号", width: 10 },
        { prop: "name", label: "姓名", width: 12 },
        { prop: "dept", label: "部门", width: 12 },
        { prop: "position", label: "职位", width: 14 },
        {
          prop: "salary",
          label: "月薪",
          width: 14,
          style: StylePresets.currency,
        },
        {
          prop: "hiredAt",
          label: "入职日期",
          width: 12,
          format: { type: "date" },
        },
      ],
      data: staffRows,
    },
  ],
});
```

## 要点

- 每个 Sheet 独立定义列、样式与数据，互不影响；
- 总行数 = 各 Sheet 之和，`auto` 模式按**总行数**路由；
- Sheet 名称遵守 Excel 规则：≤ 31 字符、不含 `: \ / ? * [ ]`。
