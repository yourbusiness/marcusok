import { describe, expect, it } from "vitest";
import {
  formatCellValue,
  generalNumber,
  splitSections,
} from "../numfmt/format";
import {
  EXCEL_BUILTIN_FORMATS,
  resolveFormatCode,
} from "../numfmt/builtin-table";

const F = (
  type: "number" | "string" | "boolean" | "error" | "formulaStr",
  value: string | null,
  code: string,
  dateSystem: "date1900" | "date1904" = "date1900",
) => formatCellValue(type, value, code, dateSystem).text;

const C = (
  type: "number" | "string" | "boolean" | "error" | "formulaStr",
  value: string | null,
  code: string,
) => formatCellValue(type, value, code, "date1900").color;

describe("splitSections", () => {
  it("拆段：引号/括号/转义内的分号不拆", () => {
    expect(splitSections('0.00;-0.00;"a;b";@')).toHaveLength(4);
    expect(splitSections("[Red]#,##0;[Blue]#,##0")).toHaveLength(2);
    expect(splitSections("yyyy")).toHaveLength(1);
  });
});

describe("General（自实现，Excel 15 位有效数字语义）", () => {
  it("整数与常规小数", () => {
    expect(F("number", "123", "General")).toBe("123");
    expect(F("number", "0.1", "General")).toBe("0.1");
  });
  it("1/3 → 15 位有效数字（引擎缺陷是 16 位尾迹）", () => {
    expect(F("number", String(1 / 3), "General")).toBe("0.333333333333333");
  });
  it("大数走科学计数近似", () => {
    expect(generalNumber(1.2345678901234568e29)).toMatch(/^1\.23457E\+29$/);
  });
  it("科学计数剥尾零且指数两位（[1e-10,1e-7) 区间的 toPrecision(15) 指数形态）", () => {
    // 此前该分支只 uppercase 不剥尾零：输出 "5.00000000000000E-8"
    expect(generalNumber(5e-8)).toBe("5E-08");
    expect(generalNumber(1.23e-9)).toBe("1.23E-09");
    expect(generalNumber(1e16)).toBe("1E+16");
    expect(generalNumber(-5e-8)).toBe("-5E-08");
  });
  it("文本/布尔/错误原样", () => {
    expect(F("string", "abc", "General")).toBe("abc");
    expect(F("boolean", "1", "General")).toBe("TRUE");
    expect(F("boolean", "0", "General")).toBe("FALSE");
    expect(F("error", "#N/A", "General")).toBe("#N/A");
  });
});

describe("数字段字面量补偿（引擎丢弃字面量的补偿层）", () => {
  it("货币前缀回注：¥/$/€", () => {
    expect(F("number", "1234.5", "¥#,##0.00")).toBe("¥1,234.50");
    expect(F("number", "1234.5", "$#,##0.00")).toBe("$1,234.50");
  });
  it("引号后缀回注：单位", () => {
    expect(F("number", "123", '0.0" kg"')).toBe("123.0 kg");
  });
  it("显式负段负号回注（Excel 语义：段内无负号则不显示）", () => {
    expect(F("number", "-1234.5", "#,##0.00;-#,##0.00")).toBe("-1,234.50");
    expect(F("number", "-1234.5", "#,##0.00;#,##0.00")).toBe("1,234.50");
  });
  it("会计括号回注", () => {
    expect(F("number", "-1234.5", "#,##0.00;(#,##0.00)")).toBe("(1,234.50)");
  });
  it("单段格式负号保留（引擎原生行为，验证未回归）", () => {
    expect(F("number", "-1234.5", "#,##0.00")).toBe("-1,234.50");
  });
  it("单段带货币前缀的负值：负号在字面量前（Excel 语义）", () => {
    expect(F("number", "-1234.5", "¥#,##0.00")).toBe("-¥1,234.50");
  });
  it("百分号/分组：交引擎", () => {
    expect(F("number", "0.5", "0%")).toBe("50%");
    expect(F("number", "1234.5", "#,##0.00")).toBe("1,234.50");
  });
  it("零段字面量原生无损（实测引擎行为，验证未回归）", () => {
    expect(F("number", "0", '0.00;-0.00;"ZERO"')).toBe("ZERO");
  });
  it("四段格式：负值选负段、文本选 @ 段", () => {
    expect(F("number", "-5", "0;0;0;@")).toBe("5"); // 无负号负段 → 无符号（Excel 同）
    expect(F("string", "abc", '"文本:"@')).toBe("文本:abc");
  });
});

describe("颜色段", () => {
  it("[Red] 提取为 CSS 色", () => {
    expect(C("number", "-1234.5", "#,##0.00;[Red]-#,##0.00")).toBe("#ff0000");
    expect(F("number", "-1234.5", "#,##0.00;[Red]-#,##0.00")).toBe("-1,234.50");
  });
  it("[Color n] indexed 色表", () => {
    expect(C("number", "1", "[Color 3]0")).toBe("#00ff00");
  });
});

describe("日期段（原样交引擎 + 1904 平移）", () => {
  it("内置 14 的 Excel 实际行为 m/d/yyyy", () => {
    expect(F("number", "45678.5", "m/d/yyyy")).toBe("1/21/2025");
  });
  it("日期字面量无损（引擎原生）", () => {
    expect(F("number", "45678.5", 'yyyy"年"m"月"d"日"')).toBe("2025年1月21日");
  });
  it("date1904：serial 0 = 1904-01-01（+1462 平移）", () => {
    expect(F("number", "0", "yyyy-m-d", "date1904")).toBe("1904-1-1");
    expect(F("number", "1.5", "yyyy-m-d h:mm", "date1904")).toBe(
      "1904-1-2 12:00",
    );
  });
  it("1900 系统下同 serial 不平移", () => {
    expect(F("number", "0", "yyyy-m-d", "date1900")).toBe("1899-12-31");
  });
  it("1900 幽灵日：serial 60 显示为 1900-02-29（Excel 保留的 Lotus 闰年 bug）", () => {
    // 引擎 serialToDate 把 60 折叠成 2/28，渲染层按 Excel 显示口径回补一天；
    // 59 / 61 两侧不受影响
    expect(F("number", "59", "m/d/yyyy")).toBe("2/28/1900");
    expect(F("number", "60", "m/d/yyyy")).toBe("2/29/1900");
    expect(F("number", "61", "m/d/yyyy")).toBe("3/1/1900");
  });
});

describe("累计时长（自实现，引擎全坏）", () => {
  it("[h]:mm:ss", () => {
    expect(F("number", "1.25", "[h]:mm:ss")).toBe("30:00:00");
    expect(F("number", "0.041666666666666664", "[h]:mm:ss")).toBe("1:00:00");
  });
  it("[m]:ss 累计分钟", () => {
    expect(F("number", "0.25", "[m]:ss")).toBe("360:00");
  });
  it("无括号 mm:ss 是当日时间的分:秒（Excel 语义；引擎误作月:秒）", () => {
    // 0.25 = 06:00:00 → "00:00"；累计分钟 "360:00" 是带括号 [m]:ss 的语义
    expect(F("number", "0.25", "mm:ss")).toBe("00:00");
    // 12:01:00 → 分钟取当日时间分量（邻接消歧：h 后的 m 是分钟）
    expect(F("number", String(0.5 + 1 / 24 / 60), "h:mm:ss")).toBe("12:01:00");
  });
  it("小数秒 mm:ss.0（内置 47 形态，四舍五入到显示精度）", () => {
    // 0.25 + 1.5s = 06:00:01.5 → "00:01.5"
    expect(F("number", String(0.25 + 1.5 / 86400), "mm:ss.0")).toBe("00:01.5");
  });
  it("[s]", () => {
    expect(F("number", "0.25", "[s]")).toBe("21600");
  });
  it("累计秒的小数秒字面量被一并消费（此前残留成 21600.0.0）", () => {
    expect(F("number", "0.25", "[ss].0")).toBe("21600.0");
    expect(F("number", "0.25", "[s].00")).toBe("21600.00");
  });
  it("[h]/[m] 段小数秒的整数位补零（此前 6:00:0.0 / 360:0.0）", () => {
    expect(F("number", "0.25", "[h]:mm:ss.0")).toBe("6:00:00.0");
    expect(F("number", String(0.25 + 7.3 / 86400), "[h]:mm:ss.0")).toBe(
      "6:00:07.3",
    );
    expect(F("number", String(0.25 + 1.5 / 86400), "[m]:ss.0")).toBe(
      "360:01.5",
    );
  });
  it('引号与转义字面量剥离（此前渲染出 5"h"04 / 5\\h\\04）', () => {
    expect(F("number", "0.2118", '[h]"h"mm"m"ss"s"')).toBe("5h04m59s");
    // \" 是字面量引号（Excel 语义），反斜杠不显示
    expect(F("number", "0.2118", '[h]\\"h\\"mm')).toBe('5"h"04');
  });
  it("负值时长显示 ####（Excel 语义，与日期段口径一致；此前按绝对值渲染）", () => {
    expect(F("number", "-0.25", "[h]:mm:ss")).toBe("#######");
    expect(F("number", "-1", "[m]:ss")).toBe("#######");
  });
});

describe("条件分段", () => {
  it("按条件选段", () => {
    expect(F("number", "95", "[Red][>=90]0;[Blue][<60]0;0")).toBe("95");
    expect(C("number", "95", "[Red][>=90]0;[Blue][<60]0;0")).toBe("#ff0000");
    expect(F("number", "50", "[Red][>=90]0;[Blue][<60]0;0")).toBe("50");
    expect(C("number", "50", "[Red][>=90]0;[Blue][<60]0;0")).toBe("#0000ff");
    expect(F("number", "85", "[Red][>=90]0;[Blue][<60]0;0")).toBe("85");
    // 无一命中时走"第一个无条件段"（Excel else 语义，且不继承条件段的颜色）：
    // 此前回落到 section[0]，85 会带 Red、5 配 [<0] 段会输出 "neg5"
    expect(C("number", "85", "[Red][>=90]0;[Blue][<60]0;0")).toBe(undefined);
    expect(F("number", "5", '[<0]"neg";[>1000]"big";0')).toBe("5");
    expect(C("number", "5", '[Red][<0]"neg";[Blue][>1000]"big";0')).toBe(
      undefined,
    );
  });
});

describe("[$display-locale] 货币/区域标记", () => {
  it("纯 locale 标记不显示（日期格式常见）", () => {
    expect(F("number", "45678.5", "[$-409]dddd, mmmm d, yyyy")).toBe(
      "Tuesday, January 21, 2025",
    );
  });
  it("货币符号按字面量显示", () => {
    expect(F("number", "1234.5", "[$$-en-US] #,##0.00")).toBe("$ 1,234.50");
    expect(F("number", "1234.5", "[$¥-804]#,##0.00")).toBe("¥1,234.50");
  });
  it("负段内同样生效（负号 + 货币符号）", () => {
    expect(F("number", "-1234.5", "[$¥-804]#,##0.00;-[$¥-804]#,##0.00")).toBe(
      "-¥1,234.50",
    );
  });
});

describe("空值与防御", () => {
  it("空值 → 空串（公式无缓存值）", () => {
    expect(F("number", null, "#,##0.00")).toBe("");
    expect(F("string", null, "General")).toBe("");
  });
  it("非法格式码不抛错（引擎兜底 String()）", () => {
    expect(() => F("number", "1", "!!![[[")).not.toThrow();
  });
  it("文本配数字格式原样显示（Excel 无 @ 段语义）", () => {
    expect(F("string", "abc", "#,##0.00")).toBe("abc");
  });
});

describe("内置行为表", () => {
  it("关键 id 的 Excel 实际串", () => {
    expect(EXCEL_BUILTIN_FORMATS.get(14)).toBe("m/d/yyyy");
    expect(EXCEL_BUILTIN_FORMATS.get(46)).toBe("[h]:mm:ss");
    expect(EXCEL_BUILTIN_FORMATS.get(49)).toBe("@");
  });
  it("resolveFormatCode：自定义优先、内置次之、兜底 General", () => {
    const custom = new Map([[164, "yyyy-mm-dd"]]);
    expect(resolveFormatCode(164, custom, () => undefined)).toBe("yyyy-mm-dd");
    expect(resolveFormatCode(3, custom, () => undefined)).toBe("#,##0");
    expect(resolveFormatCode(0, custom, () => undefined)).toBe("General");
    expect(resolveFormatCode(9999, custom, () => undefined)).toBe("General");
  });
});
