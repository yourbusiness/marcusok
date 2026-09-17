import type { Workbook } from "modern-xlsx";
import type { CellStyle } from "./types";

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
