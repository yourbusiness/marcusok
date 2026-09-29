/**
 * 样式演示的场景定义：8 种内置预设的速查元数据 + 3 个场景配方。
 *
 * 每个场景的 `sheet` 是唯一事实源——预览表格与真实导出共用同一个
 * SheetConfig 对象；`code` 是与 sheet 定义并排策展的展示文本（保留
 * StylePresets.x 的引用形态，教学价值优先），两者靠注释互指、改动需同步。
 */
import { StylePresets } from "@marcusok/excel-exporter";
import type { SheetConfig, StylePresetName } from "@marcusok/excel-exporter";
import { createDataset } from "../../mock/rows.js";

/** 演示数据：8 行、seed 固定，保证预览与每次导出的内容一致。 */
const BASE_ROWS = createDataset(8);

/** 场景①的派生字段：-0.27 ~ 0.64 的确定性增长率（千分位精度）。 */
const COMBO_ROWS = BASE_ROWS.map((r, i) => ({
  ...r,
  rate: Math.round((i * 0.13 - 0.27) * 1000) / 1000,
}));

// ---------------------------------------------------------------------------
// 内置预设速查（说明文案对齐文档站 guide/05-styles.md 的预设表）
// ---------------------------------------------------------------------------

export interface PresetCheatItem {
  name: StylePresetName;
  desc: string;
  /** 预览格的示例值：数字/日期预设给可格式化的值，其余给展示文本。 */
  sample: string | number;
}

export const PRESET_CHEATSHEET: readonly PresetCheatItem[] = [
  { name: "header", desc: "深蓝底白字加粗，12 号，居中", sample: "表头文字" },
  { name: "currency", desc: "千分位 + 两位小数，右对齐", sample: 12345.6 },
  { name: "percent", desc: "百分比两位小数，右对齐", sample: 0.234 },
  { name: "date", desc: "yyyy-MM-dd，居中", sample: "2026-07-01" },
  {
    name: "datetime",
    desc: "yyyy-MM-dd HH:mm，居中",
    sample: "2026-07-01 15:30",
  },
  { name: "dataRow", desc: "左对齐、垂直居中，底部浅灰细线", sample: "数据行" },
  {
    name: "bordered",
    desc: "四边浅灰细线，适合配表级 dataStyle",
    sample: "边框",
  },
  { name: "danger", desc: "红色加粗，居中", sample: "缺货" },
];

// ---------------------------------------------------------------------------
// 场景配方
// ---------------------------------------------------------------------------

export type ScenarioKey = "combo" | "derive" | "merge";

export interface StyleScenario {
  key: ScenarioKey;
  title: string; // Segmented 选项文案
  intro: string; // 场景说明
  filename: string;
  sheet: SheetConfig; // 预览 + 导出共用
  code: string; // 展示的 TS 代码
  notes: { text: string; tone?: "info" | "warning" }[];
  tags: string[]; // 预览卡右上角的配置标签
}

export const STYLE_SCENARIOS: readonly StyleScenario[] = [
  {
    key: "combo",
    title: "组合使用",
    intro:
      "预设为组合而设计：一个表级 headerStyle、一个表级 dataStyle，加上少量列级 style，就能产出完整的报表观感。",
    filename: "styles-combo.xlsx",
    // 下方 code 文本与本 sheet 定义一一对应，改动需同步
    sheet: {
      name: "订单月报",
      headerStyle: StylePresets.header, // 表级：整表深蓝表头
      dataStyle: StylePresets.bordered, // 表级：整表四边细线
      // 序号列表头默认「序号」、默认宽 6，用简写即可
      indexColumn: true,
      freezeRows: 1,
      autoFilter: true,
      columns: [
        {
          prop: "orderDate",
          label: "日期",
          width: 14,
          format: { type: "date" },
          style: StylePresets.date,
        },
        { prop: "name", label: "客户", width: 16, style: StylePresets.dataRow },
        { prop: "city", label: "城市", width: 12, style: StylePresets.dataRow },
        {
          prop: "amount",
          label: "金额",
          width: 14,
          style: StylePresets.currency,
        },
        {
          prop: "rate",
          label: "增长率",
          width: 12,
          style: StylePresets.percent,
        },
        {
          prop: "status",
          label: "状态",
          width: 10,
          style: StylePresets.danger,
        },
      ],
      data: COMBO_ROWS,
    },
    code: `import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "styles-combo.xlsx",
  sheets: [
    {
      name: "订单月报",
      headerStyle: StylePresets.header, // 表级：整表深蓝表头
      dataStyle: StylePresets.bordered, // 表级：整表四边细线
      indexColumn: true, // 序号列表头默认「序号」、默认宽 6
      freezeRows: 1,
      autoFilter: true,
      columns: [
        { prop: "orderDate", label: "日期", width: 14,
          format: { type: "date" }, style: StylePresets.date },
        { prop: "name", label: "客户", width: 16, style: StylePresets.dataRow },
        { prop: "city", label: "城市", width: 12, style: StylePresets.dataRow },
        { prop: "amount", label: "金额", width: 14, style: StylePresets.currency },
        { prop: "rate", label: "增长率", width: 12, style: StylePresets.percent },
        { prop: "status", label: "状态", width: 10, style: StylePresets.danger },
      ],
      data, // 8 行演示数据
    },
  ],
});`,
    notes: [
      {
        text: "金额 / 增长率列只声明 numFormat 与对齐——四边框来自表级 dataStyle，字段级深合并不会被列样式冲掉。",
      },
      {
        text: "indexColumn 的表头默认继承表级 headerStyle（文案默认「序号」），序号数据格与 dataStyle 合并；序号值由行号生成，不读 data。",
      },
      {
        text: "所有单元格默认居中：库级基底样式为未显式声明对齐的单元格补上水平 + 垂直居中，因此「客户 / 城市」这类无样式列也是居中的；预设各自声明的对齐（currency 右对齐等）优先于它。",
      },
      {
        text: "freezeRows / autoFilter 是布局特性：预览不模拟，导出后生效（冻结首行 + 表头筛选）。",
      },
      {
        text: "样式复用零成本：多个列引用同一预设，引擎对结构相同的样式去重、共享样式索引。",
      },
    ],
    tags: ["indexColumn", "freezeRows: 1", "autoFilter"],
  },
  {
    key: "derive",
    title: "spread 派生",
    intro:
      "预设是普通的常量对象（未冻结）：用对象展开从预设派生自己的变体，比从零重写 CellStyle 更简洁。",
    filename: "styles-derive.xlsx",
    // 下方 code 文本与本 sheet 定义一一对应，改动需同步
    sheet: {
      name: "派生变体",
      headerStyle: {
        ...StylePresets.header,
        fill: { pattern: "solid", fgColor: "2E7D32" },
      },
      dataStyle: {
        ...StylePresets.dataRow,
        border: StylePresets.bordered.border,
      },
      // 对象形态：自定义表头文案与列宽（默认「序号」/ 6）
      indexColumn: { label: "行号", width: 8 },
      columns: [
        { prop: "name", label: "客户", width: 16 },
        {
          prop: "amount",
          label: "金额·加粗",
          width: 16,
          style: { ...StylePresets.currency, font: { bold: true } },
        },
        {
          prop: "amount",
          label: "金额·整数千分位",
          width: 18,
          style: { ...StylePresets.currency, numFormat: "#,##0" },
        },
        {
          prop: "amount",
          label: "金额·人民币",
          width: 16,
          style: { ...StylePresets.currency, numFormat: '"¥"#,##0.00' },
        },
        {
          prop: "orderDate",
          label: "日期·中文",
          width: 18,
          format: { type: "date" },
          style: {
            numFormat: 'yyyy"年"M"月"d"日"',
            alignment: { ...StylePresets.date.alignment },
          },
        },
      ],
      data: BASE_ROWS,
    },
    code: `import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "styles-derive.xlsx",
  sheets: [
    {
      name: "派生变体",
      // 同一款表头预设换个品牌色
      headerStyle: { ...StylePresets.header,
        fill: { pattern: "solid", fgColor: "2E7D32" } },
      // dataRow 的对齐 + bordered 的四边框，组合成新的表级基底
      dataStyle: { ...StylePresets.dataRow,
        border: StylePresets.bordered.border },
      // 序号列的对象形态：自定义文案与列宽（默认「序号」/ 6）
      indexColumn: { label: "行号", width: 8 },
      columns: [
        { prop: "name", label: "客户", width: 16 },
        { prop: "amount", label: "金额·加粗", width: 16,
          style: { ...StylePresets.currency, font: { bold: true } } },
        { prop: "amount", label: "金额·整数千分位", width: 18,
          style: { ...StylePresets.currency, numFormat: "#,##0" } },
        { prop: "amount", label: "金额·人民币", width: 16,
          style: { ...StylePresets.currency, numFormat: '"¥"#,##0.00' } },
        { prop: "orderDate", label: "日期·中文", width: 18,
          format: { type: "date" },
          style: {
            numFormat: 'yyyy"年"M"月"d"日"',
            alignment: { ...StylePresets.date.alignment },
          } },
      ],
      data, // 同一 amount 字段配三列，对比不同格式
    },
  ],
});`,
    notes: [
      {
        text: "对象展开是浅合并：嵌套的 font / fill / alignment / border 会被整体替换。要保留预设的其它字段，必须手动展开内层——日期列的 alignment 就是这么写的。",
        tone: "warning",
      },
      {
        text: '格式码中的字面量文本用双引号包裹（"¥"、"年"）；yyyy"年"M"月"d"日" 里的单字符 M / d 输出不补零的月份与日期。',
      },
      {
        text: '日期列必须配 format: { type: "date" } 才会把 ISO 字符串转成 Excel 日期 serial——numFormat 只对数值单元格生效。',
      },
      {
        text: "同一 prop 可以配多列（如这里的三列金额），对比不同格式；只有 __index__ 是保留字。",
      },
    ],
    tags: ["spread 派生", "numFormat 字面量", "同 prop 多列"],
  },
  {
    key: "merge",
    title: "合并语义",
    intro:
      "两类挂载点的合并语义刻意不同：表级 dataStyle 与列级 style 是字段级深合并，表级 headerStyle 与列级 headerStyle 是整体替换。",
    filename: "styles-merge.xlsx",
    // 下方 code 文本与本 sheet 定义一一对应，改动需同步
    sheet: {
      name: "合并语义",
      headerStyle: StylePresets.header,
      dataStyle: StylePresets.bordered,
      columns: [
        { prop: "name", label: "客户·纯表级", width: 16 },
        {
          prop: "amount",
          label: "金额·深合并",
          width: 16,
          style: StylePresets.currency,
        },
        {
          prop: "status",
          label: "状态·深合并",
          width: 16,
          style: StylePresets.danger,
        },
        {
          prop: "orderDate",
          label: "表头·整体替换",
          width: 18,
          format: { type: "date" },
          headerStyle: {
            fill: { pattern: "solid", fgColor: "DDEBF7" },
            font: { bold: true, color: "1F4E79" },
          },
        },
      ],
      data: BASE_ROWS,
    },
    code: `import { exportExcel, StylePresets } from "@marcusok/excel-exporter";

await exportExcel({
  filename: "styles-merge.xlsx",
  sheets: [
    {
      name: "合并语义",
      headerStyle: StylePresets.header,
      dataStyle: StylePresets.bordered,
      columns: [
        { prop: "name", label: "客户·纯表级", width: 16 }, // 只有 dataStyle
        { prop: "amount", label: "金额·深合并", width: 16,
          style: StylePresets.currency }, // 边框保留
        { prop: "status", label: "状态·深合并", width: 16,
          style: StylePresets.danger }, // 边框保留
        { prop: "orderDate", label: "表头·整体替换", width: 18,
          format: { type: "date" }, // 自动注入 yyyy-MM-dd
          headerStyle: { // 表头整体替换：白字 / 12 号 / 居中
            fill: { pattern: "solid", fgColor: "DDEBF7" }, // 全部丢弃，
            font: { bold: true, color: "1F4E79" } } }, // 只有这里声明的生效
      ],
      data,
    },
  ],
});`,
    notes: [
      {
        text: "数据格：库级基底（默认居中）在最底层，表级 dataStyle 覆盖它，列级 style 再按字段深合并覆盖——金额 / 状态列只声明格式与字体，四边框保留。",
      },
      {
        text: "表头相反：列级 headerStyle 整体替换表级默认，不做字段合并——「表头·整体替换」列的白字、12 号、居中全部丢弃，只保留自己声明的浅蓝底与深蓝加粗字。",
        tone: "warning",
      },
      {
        text: "日期列没有列级样式时，FormatSpec 自动注入 yyyy-MM-dd；列级显式 numFormat 优先于自动注入。",
      },
    ],
    tags: ["dataStyle 深合并", "headerStyle 整体替换"],
  },
];
