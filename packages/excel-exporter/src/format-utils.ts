import type { ColumnConfig, FormatSpec, SheetConfig } from "./types";
import { dateToSerial } from "modern-xlsx";

/** Default display patterns (Excel format codes) when FormatSpec omits `pattern`. */
export const DEFAULT_DATE_PATTERN = "yyyy-MM-dd";
export const DEFAULT_DATETIME_PATTERN = "yyyy-MM-dd HH:mm";

/** Safely stringify any value to a string (objects -> JSON, null/undef -> ''). */
export function toStr(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "string") return value;
  if (value instanceof Date) return value.toISOString();
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "bigint"
  )
    return String(value);
  return JSON.stringify(value);
}

/**
 * Apply a FormatSpec to a raw value. Shared by WorkbookBuilder, StreamingBuilder,
 * and the worker entrypoint (FormatSpec is structured-clone-safe).
 */
export function applyFormat(value: unknown, spec: FormatSpec): string | number {
  switch (spec.type) {
    case "enum":
      return spec.map[toStr(value)] ?? spec.fallback ?? toStr(value);
    case "date": {
      const d = toJsDate(value);
      return d === null ? toStr(value) : dateToSerial(d);
    }
    case "datetime": {
      const d = toJsDate(value);
      return d === null ? toStr(value) : dateToSerial(d);
    }
    case "number": {
      const n = Number(value);
      if (!Number.isFinite(n)) return toStr(value);
      // Keep full precision: the stored cell value must not be truncated.
      // Display decimals/thousands are rendered via an auto-injected numFormat
      // on the workbook path (see numFormatForSpec / withAutoNumFormat). The
      // stream/SheetJS paths (no numFormat support) bake decimals into the
      // displayed value in displayValue instead.
      return n;
    }
    case "padding": {
      const s = toStr(value);
      return spec.align === "left"
        ? s.padEnd(spec.length, spec.fill)
        : s.padStart(spec.length, spec.fill);
    }
    default:
      return toStr(value);
  }
}

/**
 * Derive an Excel numFormat code from a FormatSpec so the Workbook can render
 * typed values (date serials, numbers) with the right display format. Returns
 * null for specs that produce plain strings (enum/padding) and need no numFormat.
 */
export function numFormatForSpec(spec: FormatSpec): string | null {
  switch (spec.type) {
    case "date":
      return spec.pattern ?? DEFAULT_DATE_PATTERN;
    case "datetime":
      return spec.pattern ?? DEFAULT_DATETIME_PATTERN;
    case "number": {
      const dec = spec.decimals ?? 0;
      const head = spec.thousands ? "#,##0" : "0";
      return dec > 0 ? `${head}.${"0".repeat(dec)}` : head;
    }
    default:
      return null;
  }
}

/**
 * Format a Date (or date-coercible value) into a display string using an
 * Excel-style pattern (tokens: yyyy MM dd HH mm ss). Used by the streaming
 * path, which has no numFormat support and must emit readable date strings.
 *
 * Uses the date's **UTC components** (not local ones), matching modern-xlsx's
 * `dateToSerial` (the workbook path also derives the serial from UTC
 * components). The same input therefore renders identically on the workbook,
 * stream and SheetJS paths in every timezone. Note that date-only ISO strings
 * ("2025-01-05") parse as UTC midnight per ECMA-262, while locally-constructed
 * Dates (`new Date(2025, 0, 5)`) carry local wall time whose UTC components can
 * fall on the previous day in non-UTC timezones.
 */
export function formatDateByPattern(value: unknown, pattern: string): string {
  const d = toJsDate(value);
  if (!d) return toStr(value);
  const pad = (n: number) => String(n).padStart(2, "0");
  // Excel format codes are case-insensitive, so normalize to lowercase first.
  // `mm` is ambiguous: minutes when it directly follows an hour token (`hh`),
  // otherwise the month. Scan the token stream once and resolve each `mm` from
  // its predecessor so `yyyy-mm-dd`, `yyyy-MM-dd` and `HH:mm:ss` all match.
  const lower = pattern.toLowerCase();
  const parts = {
    yyyy: String(d.getUTCFullYear()),
    month: pad(d.getUTCMonth() + 1),
    dd: pad(d.getUTCDate()),
    hh: pad(d.getUTCHours()),
    minute: pad(d.getUTCMinutes()),
    ss: pad(d.getUTCSeconds()),
  };
  const TOKEN = /yyyy|mm|dd|hh|ss/g;
  const hits: { tok: string; idx: number }[] = [];
  let mt: RegExpExecArray | null;
  while ((mt = TOKEN.exec(lower)) !== null) {
    hits.push({ tok: mt[0], idx: mt.index });
  }
  let out = "";
  let lastEnd = 0;
  for (let i = 0; i < hits.length; i++) {
    const { tok, idx } = hits[i];
    out += lower.slice(lastEnd, idx);
    lastEnd = idx + tok.length;
    if (tok === "mm") {
      // Minute only when directly preceded by an hour token; else month.
      out += hits[i - 1]?.tok === "hh" ? parts.minute : parts.month;
    } else {
      out += parts[tok as keyof typeof parts];
    }
  }
  out += lower.slice(lastEnd);
  return out;
}

/**
 * Resolve a column value to its display form: typed (number/boolean) when the
 * cell supports it, or a pattern-formatted string for dates. Shared by the
 * streaming path and the SheetJS fallback, which both lack numFormat support.
 */
export function displayValue(
  col: ColumnConfig,
  row: Record<string, unknown>,
): string | number | boolean {
  const spec = typeof col.format === "object" ? col.format : null;
  if (spec) {
    if (spec.type === "date" || spec.type === "datetime") {
      const pattern =
        spec.type === "datetime"
          ? (spec.pattern ?? DEFAULT_DATETIME_PATTERN)
          : (spec.pattern ?? DEFAULT_DATE_PATTERN);
      return formatDateByPattern(row[col.key ?? ""], pattern);
    }
    if (spec.type === "number") {
      // Stream/SheetJS paths have no numFormat support, so the configured
      // decimals must be baked into the displayed value here. The workbook
      // path keeps full precision and renders decimals via numFormat instead.
      const n = Number(row[col.key ?? ""]);
      if (!Number.isFinite(n)) return toStr(row[col.key ?? ""]);
      return Number(n.toFixed(spec.decimals ?? 0));
    }
  }
  const v = resolveCellFormat(col, row);
  // NaN/Infinity are not valid xsd:double values: writing <v>NaN</v> produces a
  // workbook Excel flags as corrupt (the same applies inside SheetJS's
  // aoa_to_sheet). Emit the visible string form instead, matching the
  // number-spec branch above.
  if (typeof v === "number") return Number.isFinite(v) ? v : toStr(v);
  if (typeof v === "boolean") return v;
  return toStr(v);
}

/**
 * Unified cell-value resolver (fixes the v1.9 format union bug): dispatches
 * function form directly, FormatSpec via applyFormat. Verified by minimal repro.
 */
export function resolveCellFormat(
  col: ColumnConfig,
  item: Record<string, unknown>,
): unknown {
  // `col.key` is optional at the type level (group columns omit it); callers
  // pass flattened leaves, whose keys are validated by flattenColumnTree.
  const raw = item[col.key ?? ""];
  if (!col.format) return raw ?? "";
  if (typeof col.format === "function") return col.format(raw, item);
  return applyFormat(raw, col.format);
}

function toJsDate(value: unknown): Date | null {
  if (value instanceof Date)
    return Number.isNaN(value.getTime()) ? null : value;
  if (typeof value === "number" || typeof value === "string") {
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  return null;
}

const SHEET_NAME_FORBIDDEN = /[\\/?*[\]:]/;

/**
 * Validate a sheet name per ECMA-376 / Excel constraints. Throws on names that
 * would produce a corrupt workbook: empty, longer than 31 chars, or containing
 * any of `: \ / ? * [ ]`.
 */
export function validateSheetName(name: string): void {
  if (typeof name !== "string" || name.length === 0) {
    throw new Error("[excel-exporter] sheet name must be a non-empty string");
  }
  if (name.length > 31) {
    throw new Error(
      `[excel-exporter] sheet name "${name.slice(0, 31)}…" exceeds the 31-char Excel limit`,
    );
  }
  if (SHEET_NAME_FORBIDDEN.test(name)) {
    throw new Error(
      `[excel-exporter] sheet name "${name}" contains forbidden characters (: \\ / ? * [ ])`,
    );
  }
}

/**
 * Validate user-supplied merge ranges against the sheet's data area. A merge
 * with a zero/negative span or an out-of-bounds endpoint would produce a
 * reversed or dangling range that Excel treats as a corrupt workbook, so
 * reject it here -- identically on the Workbook, stream and SheetJS paths --
 * with a message naming the offending merge.
 *
 * `MergeRange` is data-relative (row 0 = first data row): bounds are the leaf
 * column count and the data row count. Overlap is checked across the user
 * ranges themselves; they cannot collide with header merges because they sit
 * entirely below the header block.
 */
export function validateMerges(sheet: SheetConfig, leafCount: number): void {
  const merges = sheet.merges;
  if (!merges?.length) return;
  for (let i = 0; i < merges.length; i++) {
    const m = merges[i];
    const at = `merge #${i} {row: ${m.row}, col: ${m.col}, rowspan: ${m.rowspan}, colspan: ${m.colspan}}`;
    if (
      !Number.isInteger(m.row) ||
      !Number.isInteger(m.col) ||
      !Number.isInteger(m.rowspan) ||
      !Number.isInteger(m.colspan)
    ) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" ${at}: values must be integers`,
      );
    }
    if (m.row < 0 || m.col < 0) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" ${at}: row/col must be >= 0 (0-based, relative to the data area)`,
      );
    }
    if (m.rowspan < 1 || m.colspan < 1) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" ${at}: rowspan/colspan must be >= 1`,
      );
    }
    if (m.col + m.colspan > leafCount) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" ${at}: col ${m.col} + colspan ${m.colspan} exceeds the ${leafCount} leaf columns`,
      );
    }
    if (m.row + m.rowspan > sheet.data.length) {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" ${at}: row ${m.row} + rowspan ${m.rowspan} exceeds the ${sheet.data.length} data rows`,
      );
    }
    for (let j = 0; j < i; j++) {
      const o = merges[j];
      const overlaps =
        m.row < o.row + o.rowspan &&
        o.row < m.row + m.rowspan &&
        m.col < o.col + o.colspan &&
        o.col < m.col + m.colspan;
      if (overlaps) {
        throw new Error(
          `[excel-exporter] sheet "${sheet.name}" ${at} overlaps merge #${j}; merged ranges must be disjoint`,
        );
      }
    }
  }
}
