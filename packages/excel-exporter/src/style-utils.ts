import type { Workbook } from "modern-xlsx";
import type { CellStyle } from "./types";

/**
 * 库级基底样式：铺在**所有**单元格（表头格与数据格）的样式解析链最底层。
 * 原先无样式单元格遵循 Excel 的原生对齐（文本左、数字右），一张表里两种
 * 走向显得参差；铺上这层后默认水平 + 垂直居中。
 *
 * 显式声明始终优先：mergeStyles 是字段级合并，表级 dataStyle / headerStyle
 * 或列级 style / headerStyle 只要声明了 alignment（哪怕只声明 horizontal），
 * 就会覆盖对应字段，未声明的字段仍由本基底兜底。因此想要 Excel 原生左对齐
 * 的调用方，显式写 `dataStyle: { alignment: { horizontal: "left" } }` 即可。
 *
 * 与 dataStyle 的区别：dataStyle 是调用方可配可覆盖的「业务基底」，本常量是
 * 库的默认底线；两者都在最底层，dataStyle 覆盖它。
 *
 * 冻结：本常量是导出的公共对象，而 mergeStyles(base, undefined) 会**原样返回
 * base 引用**——不冻结的话，调用方一次就地修改就会改变所有后续导出的观感。
 */
export const BaseCellStyle: CellStyle = Object.freeze({
  alignment: Object.freeze({ horizontal: "center", vertical: "center" }),
});

/**
 * Field-level deep merge of two cell styles: `override`'s set fields win, the
 * rest are inherited from `base`. Serves `SheetConfig.dataStyle` — a table-wide
 * border survives a column that only sets `numFormat`, and vice versa. This is
 * deliberately different from `headerStyle`, which replaces wholesale.
 */
export function mergeStyles(
  base: CellStyle | undefined,
  override: CellStyle | undefined,
): CellStyle | undefined {
  if (!base) return override;
  if (!override) return base;
  const merged: CellStyle = {};
  if (base.font || override.font)
    merged.font = { ...base.font, ...override.font };
  if (base.fill || override.fill)
    merged.fill = { ...base.fill, ...override.fill };
  if (base.alignment || override.alignment)
    merged.alignment = { ...base.alignment, ...override.alignment };
  if (base.border || override.border)
    merged.border = { ...base.border, ...override.border };
  // numFormat 无子字段，直接择一覆盖
  if (override.numFormat !== undefined) merged.numFormat = override.numFormat;
  else if (base.numFormat !== undefined) merged.numFormat = base.numFormat;
  return merged;
}

/**
 * Compile a business CellStyle into a modern-xlsx styleIndex (0-based index into
 * the workbook's cellXfs table). StyleBuilder chain methods mutate in place and
 * return `this` (verified in modern-xlsx 1.2.0 source), so direct calls suffice.
 */
export function buildStyleIndex(wb: Workbook, style: CellStyle): number {
  const builder = wb.createStyle();

  if (style.font) {
    const { bold, italic, size, color, name } = style.font;
    builder.font({
      ...(bold !== undefined && { bold }),
      ...(italic !== undefined && { italic }),
      ...(size !== undefined && { size }),
      ...(color !== undefined && { color }),
      ...(name !== undefined && { name }),
    });
  }

  if (style.fill && (style.fill.fgColor || style.fill.bgColor)) {
    builder.fill({
      pattern: style.fill.pattern ?? "solid",
      fgColor: style.fill.fgColor ?? null,
      bgColor: style.fill.bgColor ?? null,
    });
  }

  if (style.alignment) {
    const { horizontal, vertical, wrapText, textRotation } = style.alignment;
    builder.alignment({
      ...(horizontal && { horizontal }),
      ...(vertical && { vertical }),
      ...(wrapText !== undefined && { wrapText }),
      ...(textRotation !== undefined && { textRotation }),
    });
  }

  if (style.border) {
    const { top, bottom, left, right } = style.border;
    // Skip the call when no side is defined: passing an all-undefined object
    // would rely on modern-xlsx tolerating empty border specs (unverified).
    if (top || bottom || left || right) {
      builder.border({ top, bottom, left, right });
    }
  }

  if (style.numFormat) {
    builder.numberFormat(style.numFormat);
  }

  return builder.build(wb.styles);
}
