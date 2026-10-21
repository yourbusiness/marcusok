/**
 * CellStyle 的运行时校验。
 *
 * TypeScript 只能约束 TS 调用方；JS 调用方（或 any 绕过）把 "middle"、
 * "bogus" 之类的值直接交给引擎时，modern-xlsx 会原样写进 styles.xml，
 * Excel 打开时要求修复。这里与 border.style / format.type 的既有前置换口径
 * 保持一致：枚举/类型错误在导出前失败，而不是产出坏文件或整份降级。
 *
 * 同时也被低层 WorkbookBuilder 复用（buildStyleIndex 入口），保证绕过
 * exportExcel 预检的直接调用同样不会写出非法 XML。
 */
import type { CellStyle } from "./types";
import { assertXmlSafeText } from "./xml-safety";

const COLOR_HEX = /^[0-9A-Fa-f]{6}$/;
const BORDER_STYLES = new Set<string>([
  "thin",
  "medium",
  "thick",
  "dashed",
  "dotted",
  "double",
  "hair",
  "mediumDashed",
  "dashDot",
  "mediumDashDot",
  "dashDotDot",
  "mediumDashDotDot",
  "slantDashDot",
]);
const ALIGN_HORIZONTAL = new Set(["left", "center", "right"]);
const ALIGN_VERTICAL = new Set(["top", "center", "bottom"]);
const FILL_PATTERNS = new Set(["solid", "none"]);

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function assertObject(value: unknown, at: string, field: string): void {
  if (!isPlainObject(value)) {
    throw new Error(`[excel-exporter] ${at}: ${field} must be an object`);
  }
}

function validateColor(
  color: string | undefined,
  at: string,
  field: string,
): void {
  if (color === undefined) return;
  if (typeof color !== "string" || !COLOR_HEX.test(color)) {
    throw new Error(
      `[excel-exporter] ${at}: ${field} must be a 6-digit RGB hex string (e.g. 'FF0000'), got ${JSON.stringify(color)}`,
    );
  }
}

/** 校验一个 CellStyle；`at` 用于把错误定位到具体表/列/表头。 */
export function validateCellStyle(
  style: CellStyle | undefined,
  at: string,
): void {
  if (style === undefined || style === null) return;
  assertObject(style, at, "style");

  const font = style.font as unknown;
  if (font !== undefined) {
    assertObject(font, at, "font");
    const f = font as NonNullable<CellStyle["font"]>;
    if (f.bold !== undefined && typeof f.bold !== "boolean") {
      throw new Error(`[excel-exporter] ${at}: font.bold must be a boolean`);
    }
    if (f.italic !== undefined && typeof f.italic !== "boolean") {
      throw new Error(`[excel-exporter] ${at}: font.italic must be a boolean`);
    }
    if (f.name !== undefined) assertXmlSafeText(f.name, at, "font.name");
    const size = f.size;
    if (
      size !== undefined &&
      (typeof size !== "number" || !Number.isFinite(size) || size <= 0)
    ) {
      throw new Error(
        `[excel-exporter] ${at}: font.size must be a finite positive number`,
      );
    }
    validateColor(f.color, at, "font.color");
  }

  const fill = style.fill as unknown;
  if (fill !== undefined) {
    assertObject(fill, at, "fill");
    const f = fill as NonNullable<CellStyle["fill"]>;
    if (f.pattern !== undefined && !FILL_PATTERNS.has(f.pattern)) {
      throw new Error(
        `[excel-exporter] ${at}: fill.pattern must be "solid" or "none" (got ${JSON.stringify(f.pattern)})`,
      );
    }
    validateColor(f.fgColor, at, "fill.fgColor");
    validateColor(f.bgColor, at, "fill.bgColor");
  }

  const alignment = style.alignment as unknown;
  if (alignment !== undefined) {
    assertObject(alignment, at, "alignment");
    const a = alignment as NonNullable<CellStyle["alignment"]>;
    if (a.horizontal !== undefined && !ALIGN_HORIZONTAL.has(a.horizontal)) {
      throw new Error(
        `[excel-exporter] ${at}: alignment.horizontal must be "left", "center" or "right" (got ${JSON.stringify(a.horizontal)})`,
      );
    }
    if (a.vertical !== undefined && !ALIGN_VERTICAL.has(a.vertical)) {
      throw new Error(
        `[excel-exporter] ${at}: alignment.vertical must be "top", "center" or "bottom" (got ${JSON.stringify(a.vertical)})`,
      );
    }
    if (a.wrapText !== undefined && typeof a.wrapText !== "boolean") {
      throw new Error(
        `[excel-exporter] ${at}: alignment.wrapText must be a boolean`,
      );
    }
    const rotation = a.textRotation;
    if (
      rotation !== undefined &&
      (!Number.isInteger(rotation) || rotation < 0 || rotation > 180)
    ) {
      throw new Error(
        `[excel-exporter] ${at}: alignment.textRotation must be an integer between 0 and 180`,
      );
    }
  }

  const border = style.border as unknown;
  if (border !== undefined) {
    assertObject(border, at, "border");
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const b = (border as NonNullable<CellStyle["border"]>)[side];
      if (!b) continue;
      assertObject(b, at, `border.${side}`);
      if (!BORDER_STYLES.has(b.style)) {
        throw new Error(
          `[excel-exporter] ${at}: border.${side}.style must be one of the OOXML border styles (thin/medium/thick/dashed/...), got ${JSON.stringify(b.style)}`,
        );
      }
      validateColor(b.color, at, `border.${side}.color`);
    }
  }

  if (style.numFormat !== undefined) {
    assertXmlSafeText(style.numFormat, at, "numFormat");
  }
}
