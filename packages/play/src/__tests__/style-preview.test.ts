import { describe, expect, it } from "vitest";
import {
  buildPreviewColumns,
  cellStyleToCss,
  effectiveDataStyle,
  effectiveHeaderStyle,
  expandIndexColumn,
  formatByNumFormat,
  isDateNumFormat,
  previewCellValue,
} from "../demos/excel-exporter/style-preview";

// 样本样式内联字面量（与 StylePresets 等价）：测试在 node 环境运行，不能从
// 公共入口导入运行时值——那会把 dist 全量引擎拉进测试进程。
const HEADER_STYLE = {
  font: { bold: true, size: 12, color: "FFFFFF" },
  fill: { pattern: "solid", fgColor: "1F4E79" },
  alignment: { horizontal: "center", vertical: "center" },
} as const;
const BORDERED_STYLE = {
  border: {
    top: { style: "thin", color: "D0D0D0" },
    bottom: { style: "thin", color: "D0D0D0" },
    left: { style: "thin", color: "D0D0D0" },
    right: { style: "thin", color: "D0D0D0" },
  },
} as const;
const CURRENCY_STYLE = {
  numFormat: "#,##0.00",
  alignment: { horizontal: "right" },
} as const;
// 库级基底（BaseCellStyle）的对齐：钉住行为——未显式声明的对齐一律居中
const CENTER = { horizontal: "center", vertical: "center" } as const;

describe("cellStyleToCss", () => {
  it("maps every declared field of a full style", () => {
    expect(
      cellStyleToCss({
        font: { bold: true, size: 12, color: "FFFFFF", name: "Arial" },
        fill: { pattern: "solid", fgColor: "1F4E79" },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border: { bottom: { style: "thin", color: "D0D0D0" } },
      }),
    ).toEqual({
      fontWeight: 700,
      fontSize: "12pt",
      color: "#FFFFFF",
      fontFamily: "Arial",
      background: "#1F4E79",
      textAlign: "center",
      verticalAlign: "middle",
      whiteSpace: "pre-wrap",
      borderBottom: "1px solid #D0D0D0",
    });
  });

  it("returns {} for undefined and skips pattern:none fills", () => {
    expect(cellStyleToCss(undefined)).toEqual({});
    expect(
      cellStyleToCss({ fill: { pattern: "none", fgColor: "FF0000" } }),
    ).toEqual({});
  });

  it("maps every border weight tier", () => {
    expect(cellStyleToCss({ border: { top: { style: "hair" } } })).toEqual({
      borderTop: "0.5px solid #000000", // 缺省色 = Excel 自动色（黑）
    });
    expect(cellStyleToCss({ border: { top: { style: "thin" } } })).toEqual({
      borderTop: "1px solid #000000",
    });
    expect(cellStyleToCss({ border: { top: { style: "medium" } } })).toEqual({
      borderTop: "2px solid #000000",
    });
    expect(cellStyleToCss({ border: { top: { style: "thick" } } })).toEqual({
      borderTop: "3px solid #000000",
    });
    expect(cellStyleToCss({ border: { top: { style: "double" } } })).toEqual({
      borderTop: "3px double #000000",
    });
    // CSS 无点划线：dashDot 系降级 dashed
    expect(cellStyleToCss({ border: { top: { style: "dashDot" } } })).toEqual({
      borderTop: "1px dashed #000000",
    });
  });

  it("does not map textRotation (CSS rotation would distort the anchor)", () => {
    expect(cellStyleToCss({ alignment: { textRotation: 45 } })).toEqual({});
  });
});

describe("formatByNumFormat", () => {
  it("formats numbers: thousands, decimals, currency literal, percent", () => {
    expect(formatByNumFormat(12345.6, "#,##0.00")).toBe("12,345.60");
    expect(formatByNumFormat(12345.6, "#,##0")).toBe("12,346");
    expect(formatByNumFormat(-1234.5, "#,##0")).toBe("-1,235");
    expect(formatByNumFormat(12345.6, '"¥"#,##0.00')).toBe("¥12,345.60");
    expect(formatByNumFormat(0.234, "0.00%")).toBe("23.40%");
    expect(formatByNumFormat(12345.6, "0.00")).toBe("12345.60"); // 无千分位
  });

  it("formats dates from ISO strings, including CJK literal pattern", () => {
    expect(formatByNumFormat("2024-03-05", "yyyy-MM-dd")).toBe("2024-03-05");
    expect(formatByNumFormat("2026-07-01 15:30", "yyyy-MM-dd HH:mm")).toBe(
      "2026-07-01 15:30",
    );
    expect(formatByNumFormat("2024-03-05", 'yyyy"年"M"月"d"日"')).toBe(
      "2024年3月5日",
    );
  });

  it("falls back to String for unknown codes and text values under numeric codes", () => {
    expect(formatByNumFormat(123, "??unknown??")).toBe("123");
    // Excel 语义：numFormat 对文本单元格不生效
    expect(formatByNumFormat("abc", "#,##0.00")).toBe("abc");
    // 日期码遇到非 ISO 值：原样返回
    expect(formatByNumFormat("not-a-date", "yyyy-MM-dd")).toBe("not-a-date");
  });

  it("isDateNumFormat separates date codes from numeric codes", () => {
    expect(isDateNumFormat("yyyy-MM-dd")).toBe(true);
    expect(isDateNumFormat('yyyy"年"M"月"d"日"')).toBe(true);
    expect(isDateNumFormat("#,##0.00")).toBe(false);
    expect(isDateNumFormat("0.00%")).toBe(false);
    // [...] 条件/颜色段不算格式 token：'[Red]' 里的 d 不得把数字格式误判成日期
    expect(isDateNumFormat("#,##0;[Red]-#,##0")).toBe(false);
    expect(formatByNumFormat(12345.6, "#,##0;[Red]-#,##0")).toBe("12,346");
  });
});

describe("effective styles mirror the export engine semantics", () => {
  it("header: column-level headerStyle replaces the sheet default wholesale", () => {
    const colStyle = { font: { italic: true } };
    const col = { prop: "a", label: "A", headerStyle: colStyle };
    const sheet = {
      name: "S",
      columns: [col],
      data: [],
      headerStyle: HEADER_STYLE,
    };
    // 整体替换：列级对象与表级默认不合并，表级的字体/填充全被丢弃；
    // 基底铺在结果之下，两者都没声明的对齐由基底补上居中
    expect(effectiveHeaderStyle(sheet, col)).toEqual({
      font: { italic: true },
      alignment: { ...CENTER },
    });
    const plain = { prop: "b", label: "B" };
    // 表级默认整体生效（该预设自带居中，与基底取值一致）
    expect(effectiveHeaderStyle(sheet, plain)).toEqual(HEADER_STYLE);
  });

  it("lays the library base style under unstyled cells (default centering)", () => {
    const sheet = { name: "S", columns: [{ prop: "v", label: "V" }], data: [] };
    const col = sheet.columns[0]!;
    expect(effectiveDataStyle(sheet, col)).toEqual({
      alignment: { ...CENTER },
    });
    expect(effectiveHeaderStyle(sheet, col)).toEqual({
      alignment: { ...CENTER },
    });
  });

  it("data: column style deep-merges over dataStyle, field by field", () => {
    const col = { prop: "amount", label: "Amount", style: CURRENCY_STYLE };
    const sheet = {
      name: "S",
      columns: [col],
      data: [],
      dataStyle: BORDERED_STYLE,
    };
    const merged = effectiveDataStyle(sheet, col);
    // 深合并：currency 只声明 numFormat 与水平对齐，bordered 的四边框保留，
    // 垂直对齐由基底补上
    expect(merged).toEqual({
      numFormat: "#,##0.00",
      alignment: { horizontal: "right", vertical: "center" },
      border: { ...BORDERED_STYLE.border },
    });
  });

  it("data: FormatSpec auto-injects numFormat when the column style has none", () => {
    const col = { prop: "d", label: "D", format: { type: "date" } as const };
    const sheet = { name: "S", columns: [col], data: [] };
    expect(effectiveDataStyle(sheet, col)).toEqual({
      numFormat: "yyyy-MM-dd",
      alignment: { ...CENTER },
    });
    // 列级显式 numFormat 优先，注入不发生
    const explicit = {
      prop: "d2",
      label: "D2",
      format: { type: "date" } as const,
      style: { numFormat: 'yyyy"年"M"月"d"日"' },
    };
    expect(effectiveDataStyle(sheet, explicit)).toEqual({
      numFormat: 'yyyy"年"M"月"d"日"',
      alignment: { ...CENTER },
    });
  });
});

describe("index column expansion and cell values", () => {
  it("expands indexColumn with defaults: 序号 / width 6 / values 1..n", () => {
    const sheet = {
      name: "S",
      indexColumn: true,
      columns: [{ prop: "name", label: "Name" }],
      data: [{ name: "a" }, { name: "b" }],
    };
    const cols = expandIndexColumn(sheet);
    expect(cols).toHaveLength(2);
    expect(cols[0]).toMatchObject({
      prop: "__index__",
      label: "序号",
      width: 6,
    });

    const preview = buildPreviewColumns(sheet);
    expect(preview[0]!.isIndex).toBe(true);
    expect(previewCellValue(sheet, preview[0]!, sheet.data[0]!, 0)).toBe("1");
    expect(previewCellValue(sheet, preview[0]!, sheet.data[1]!, 1)).toBe("2");
    // 自定义 start
    const offset = { ...sheet, indexColumn: { start: 100 } };
    expect(
      previewCellValue(offset, buildPreviewColumns(offset)[0]!, {}, 2),
    ).toBe("102");
  });

  it("index column styles follow sheet-level semantics", () => {
    const sheet = {
      name: "S",
      indexColumn: {
        headerStyle: { font: { bold: true } }, // 覆盖表级表头样式
        style: { alignment: { horizontal: "center" } as const }, // 与 dataStyle 深合并
      },
      columns: [{ prop: "name", label: "Name" }],
      data: [],
      headerStyle: HEADER_STYLE,
      dataStyle: BORDERED_STYLE,
    };
    const preview = buildPreviewColumns(sheet);
    // 表头：整体替换表级默认，未声明的对齐由基底补上
    expect(preview[0]!.headerStyle).toEqual({
      font: { bold: true },
      alignment: { ...CENTER },
    });
    // 数据：与 dataStyle 字段级合并，边框保留
    expect(preview[0]!.dataStyle).toEqual({
      alignment: { ...CENTER },
      border: { ...BORDERED_STYLE.border },
    });
  });

  it("gives repeated props a distinct React key while keeping prop as the field", () => {
    // 派生场景刻意让三列共用 prop "amount"：id 必须唯一（否则 React key 冲突），
    // 而 prop 仍是读值的字段名
    const sheet = {
      name: "S",
      columns: [
        { prop: "amount", label: "A1" },
        { prop: "amount", label: "A2" },
        { prop: "amount", label: "A3" },
      ],
      data: [{ amount: 1 }],
    };
    const cols = buildPreviewColumns(sheet);
    expect(new Set(cols.map((c) => c.id)).size).toBe(3);
    expect(cols.every((c) => c.prop === "amount")).toBe(true);
    // 同 prop 的三列都取到同一个值
    expect(
      cols.map((c) => previewCellValue(sheet, c, sheet.data[0]!, 0)),
    ).toEqual(["1", "1", "1"]);
  });

  it("renders null/undefined as empty cells and plain values as strings", () => {
    const sheet = {
      name: "S",
      columns: [{ prop: "v", label: "V" }],
      data: [{ v: null }],
    };
    const col = buildPreviewColumns(sheet)[0]!;
    expect(previewCellValue(sheet, col, { v: null }, 0)).toBe("");
    expect(previewCellValue(sheet, col, { v: undefined }, 0)).toBe("");
    expect(previewCellValue(sheet, col, { v: "文本" }, 0)).toBe("文本");
  });
});
