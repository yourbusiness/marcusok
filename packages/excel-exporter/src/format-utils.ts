import type { ColumnConfig, FormatSpec, SheetConfig } from "./types";
import { columnProp } from "./column-tree";
import { dateToSerial } from "@marcusok/xlsx-core";
import {
  firstInvalidXmlCode,
  formatCodePoint,
  sanitizeXmlText,
} from "./xml-safety";

/** Default display patterns (Excel format codes) when FormatSpec omits `pattern`. */
export const DEFAULT_DATE_PATTERN = "yyyy-MM-dd";
export const DEFAULT_DATETIME_PATTERN = "yyyy-MM-dd HH:mm";

/** Safely stringify any value to a string (objects -> JSON, null/undef -> '').
 *  Symbols/functions are not JSON-serializable (JSON.stringify returns
 *  undefined for them); String() them so the cell never receives a non-string.
 *  Invalid Dates stringify as "Invalid Date" — toISOString() would throw a
 *  RangeError and fail the whole export over one bad cell (see toJsDate for
 *  the same NaN guard on the parsing side). */
export function toStr(value: unknown): string {
  if (value == null) return "";
  // 数据文本统一清洗 XML 非法控制字符/孤立代理项：modern-xlsx 的 JSON
  // 序列化遇到孤立代理项会直接拒绝（整份导出降级为无样式 stream），而
  // fast-xlsx 的 TextEncoder 会替换为 U+FFFD。在这里归一后两条路径写入
  // 相同内容，也避免一个坏字符触发整份降级。
  if (typeof value === "string") return sanitizeXmlText(value);
  if (value instanceof Date)
    return Number.isNaN(value.getTime())
      ? sanitizeXmlText(String(value))
      : value.toISOString();
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "bigint"
  )
    return String(value);
  if (typeof value === "symbol" || typeof value === "function")
    return sanitizeXmlText(String(value));
  // toJSON 可能返回 undefined（不少第三方对象的惯例），JSON.stringify 随之
  // 返回 undefined，与本函数的 string 契约不符：若放行，undefined 直入 aoa，
  // modern-xlsx 会跳过该格不建单元格，破坏"数据行稠密"假设——同行右侧列
  // 的样式错位一格，值也静默丢失。循环引用对象（ORM 实体互指等）则会让
  // JSON.stringify 直接抛 TypeError。两种病态都兜底成可见字符串——一个坏
  // 对象不丢整格、更不能弄挂整次导出；兜底意图就是默认字符串化
  // （"[object Object]"也远好于静默丢值）。
  try {
    // eslint-disable-next-line @typescript-eslint/no-base-to-string
    return sanitizeXmlText(JSON.stringify(value) ?? String(value));
  } catch {
    // eslint-disable-next-line @typescript-eslint/no-base-to-string
    return sanitizeXmlText(String(value));
  }
}

/**
 * 空串/纯空白串在 number 列视同缺失。这是数据库、表单、CSV 导入里最常见的
 * 缺失值形态，而 `Number("") === 0`（`Number(" ")` 同）会把它静默变成 0——
 * 0 在财务数据里是有意义的值。与 null/undefined 同等对待（渲染空单元格）。
 */
function isBlankString(value: unknown): boolean {
  return typeof value === "string" && value.trim() === "";
}

/**
 * Apply a FormatSpec to a raw value. Shared by WorkbookBuilder, StreamingBuilder,
 * and the worker entrypoint (FormatSpec is structured-clone-safe).
 */
export function applyFormat(value: unknown, spec: FormatSpec): string | number {
  switch (spec.type) {
    case "enum": {
      // 缺失值（null/undefined/空串/纯空白串）不是"未映射值"：与 number/
      // padding 分支同一判空口径，渲染空单元格，不落 fallback——否则数据
      // 空洞会被填成 fallback 文案（"未知"这类有业务含义的假值）。
      if (value == null || isBlankString(value)) return "";
      // Own-property lookup: a plain `spec.map[key]` walks the prototype
      // chain, so values like "constructor" / "toString" / "__proto__" that
      // the map does not define would find Object.prototype members (non-
      // nullish) and bypass the fallback, writing "function Object() {...}"
      // into the cell instead of the configured fallback text.
      const key = toStr(value);
      const mapped = Object.hasOwn(spec.map, key) ? spec.map[key] : undefined;
      return mapped ?? spec.fallback ?? key;
    }
    case "date": {
      // 与 number/padding/stream 路径同一判空口径：空串/纯空白串是缺失值，
      // 渲染空单元格；否则会写成带空白的文本格，与跨 50k 阈值的 stream
      // 路径（直接跳过）产生"单元格存在性"差异。
      if (value == null || isBlankString(value)) return "";
      const d = toJsDate(value);
      return d === null ? toStr(value) : dateToSerial(d);
    }
    case "datetime": {
      if (value == null || isBlankString(value)) return "";
      const d = toJsDate(value);
      return d === null ? toStr(value) : dateToSerial(d);
    }
    case "number": {
      // null/undefined (and blank strings — see isBlankString) render as an
      // empty cell on every path. Without this guard Number(null) === 0 would
      // silently turn missing values into 0 while undefined became ""
      // (asymmetric, and 0 is a meaningful value in financial data).
      if (value == null || isBlankString(value)) return "";
      const n = Number(value);
      if (!Number.isFinite(n)) return toStr(value);
      // Keep full precision: the stored cell value must not be truncated.
      // Display decimals/thousands are rendered via an auto-injected numFormat
      // on the workbook path (see numFormatForSpec / withAutoNumFormat). The
      // stream path (no numFormat support) bakes decimals into the
      // displayed value in displayValue instead.
      return n;
    }
    case "padding": {
      // 缺失值不填充：null/undefined 与空串/纯空白串（isBlankString，与
      // number 分支同一判定）都会被 padStart 填成 "00000" 这类有业务含义
      // 的假编号——与 number 分支的缺失守卫对称。
      if (value == null || isBlankString(value)) return "";
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
 * Not a full Excel format-code engine: the pattern is lower-cased before the
 * scan and quoted literals are not interpreted (their quotes are emitted
 * as-is), so `yyyy"年"M"月"d"日"` renders as `2026"年"m"月"d"日"` here, while
 * the Workbook path hands the same pattern to Excel's numFormat and renders
 * `2026年7月1日` (see the FormatSpec note in types.ts).
 *
 * Uses the date's **UTC components** (not local ones), matching modern-xlsx's
 * `dateToSerial` (the workbook path also derives the serial from UTC
 * components). The same input therefore renders identically on the workbook
 * and stream paths in every timezone. Note that date-only ISO strings
 * ("2025-01-05") parse as UTC midnight per ECMA-262, while locally-constructed
 * Dates (`new Date(2025, 0, 5)`) carry local wall time whose UTC components can
 * fall on the previous day in non-UTC timezones.
 */
export function formatDateByPattern(value: unknown, pattern: string): string {
  const d = toJsDate(value);
  if (!d) return toStr(value);
  const pad = (n: number) => String(n).padStart(2, "0");
  // Excel format codes are case-insensitive, so normalize to lowercase first.
  // 整串小写也作用于引号字面量（"T" 会渲染成 t）——见函数头注释。
  // `mm` is ambiguous: minutes when adjacent to a time token (directly after an
  // hour token `hh`, or directly before a seconds token `ss` — Excel's own
  // numFormat convention), otherwise the month. Scan the token stream once and
  // resolve each `mm` from its neighbors so `yyyy-mm-dd`, `yyyy-MM-dd`,
  // `HH:mm:ss` and `mm:ss` all match the workbook path's Excel rendering.
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
      // Minute when the previous token is an hour or the next is seconds;
      // else month.
      out +=
        hits[i - 1]?.tok === "hh" || hits[i + 1]?.tok === "ss"
          ? parts.minute
          : parts.month;
    } else {
      out += parts[tok as keyof typeof parts];
    }
  }
  out += lower.slice(lastEnd);
  return out;
}

/**
 * Number format code a column's FormatSpec needs on the stream path, or null
 * when the column produces text (enum/padding/function/no format). The
 * fast-xlsx writer emits a minimal styles.xml for these codes, so dates stay
 * real serial numbers and decimals/thousands render exactly like the
 * Workbook path instead of being baked into the stored value.
 */
export function columnNumFmt(col: ColumnConfig): string | null {
  const spec = typeof col.format === "object" ? col.format : null;
  return spec ? numFormatForSpec(spec) : null;
}

/** Normalize a resolved raw value to a stream cell value (same rules as the
 * Workbook path: finite numbers stay numeric, booleans stay booleans,
 * everything else becomes its visible string form). */
function normalizeStreamValue(value: unknown): string | number | boolean {
  if (typeof value === "number")
    return Number.isFinite(value) ? value : toStr(value);
  if (typeof value === "boolean") return value;
  return toStr(value);
}

/**
 * Streaming-path cell resolver. Returns the **stored** value (not a display
 * string) so fast-xlsx can pair it with a numFmt style:
 * - date/datetime: Excel serial number (UTC components, same `dateToSerial`
 *   as the Workbook path); invalid dates fall back to their visible string.
 * - number: full-precision number (decimals/thousands live in the numFmt).
 * - enum/padding/function/no format: the same normalized value as before.
 * Empty results are returned as `""`; the writer skips them (no cell).
 *
 * `displayValue` is kept unchanged for public callers that need the rendered
 * string form; this function is the storage-level counterpart used by
 * fast-xlsx.
 */
export function streamCellValue(
  col: ColumnConfig,
  row: Record<string, unknown>,
): string | number | boolean {
  const spec = typeof col.format === "object" ? col.format : null;
  if (spec) {
    if (spec.type === "date" || spec.type === "datetime") {
      const raw = row == null ? undefined : row[columnProp(col) ?? ""];
      if (raw == null || isBlankString(raw)) return "";
      const d = toJsDate(raw);
      return d === null ? toStr(raw) : dateToSerial(d);
    }
    if (spec.type === "number") {
      const raw = row == null ? undefined : row[columnProp(col) ?? ""];
      if (raw == null || isBlankString(raw)) return "";
      const n = Number(raw);
      return Number.isFinite(n) ? n : toStr(raw);
    }
    return normalizeStreamValue(
      applyFormat(row == null ? undefined : row[columnProp(col) ?? ""], spec),
    );
  }
  return normalizeStreamValue(resolveCellFormat(col, row));
}

/**
 * Resolve a column value to its display form: typed (number/boolean) when the
 * cell supports it, or a pattern-formatted string for dates. Used by public
 * callers that need rendered text; fast-xlsx uses {@link streamCellValue}
 * (storage-level values + numFmt) instead.
 */
export function displayValue(
  col: ColumnConfig,
  row: Record<string, unknown>,
): string | number | boolean {
  const spec = typeof col.format === "object" ? col.format : null;
  if (spec) {
    if (spec.type === "date" || spec.type === "datetime") {
      const raw = row == null ? undefined : row[columnProp(col) ?? ""];
      // 与 applyFormat / streamCellValue 对齐：空串/纯空白串是缺失值。
      if (raw == null || isBlankString(raw)) return "";
      const pattern =
        spec.type === "datetime"
          ? (spec.pattern ?? DEFAULT_DATETIME_PATTERN)
          : (spec.pattern ?? DEFAULT_DATE_PATTERN);
      return formatDateByPattern(raw, pattern);
    }
    if (spec.type === "number") {
      // The stream path has no numFormat support, so the configured
      // decimals must be baked into the displayed value here. The workbook
      // path keeps full precision and renders decimals via numFormat instead.
      // null/undefined/blank strings render as an empty cell, mirroring
      // applyFormat (never Number(null) === 0 — see isBlankString there).
      const raw = row == null ? undefined : row[columnProp(col) ?? ""];
      if (raw == null || isBlankString(raw)) return "";
      const n = Number(raw);
      if (!Number.isFinite(n)) return toStr(raw);
      return Number(n.toFixed(spec.decimals ?? 0));
    }
  }
  const v = resolveCellFormat(col, row);
  // NaN/Infinity are not valid xsd:double values: writing <v>NaN</v> produces a
  // workbook Excel flags as corrupt. Emit the visible string form instead,
  // matching the number-spec branch above.
  if (typeof v === "number") return Number.isFinite(v) ? v : toStr(v);
  if (typeof v === "boolean") return v;
  return toStr(v);
}

/**
 * Unified cell-value resolver: dispatches the function form directly and the
 * FormatSpec form via applyFormat. (The pre-1.1 bug this replaced dispatched
 * only one arm of the union, so FormatSpec columns silently fell through to
 * raw values.)
 */
export function resolveCellFormat(
  col: ColumnConfig,
  item: Record<string, unknown>,
): unknown {
  // `col.prop` is optional at the type level (group columns omit it); callers
  // pass flattened leaves, whose props are validated by flattenColumnTree.
  // A null/primitive row (pathological input) reads as all-fields-missing
  // instead of throwing a raw TypeError on property access.
  const raw = item == null ? undefined : item[columnProp(col) ?? ""];
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
 * would produce a corrupt workbook: empty, longer than 31 chars, containing
 * any of `: \ / ? * [ ]`, or beginning/ending with an apostrophe (Excel rejects
 * such names even though the character is legal inside the name).
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
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new Error(
      `[excel-exporter] sheet name "${name}" must not begin or end with an apostrophe (')`,
    );
  }
  // XML 1.0 不允许的码点会被现代写 API 原样拼进 workbook.xml，Excel 打开
  // 时要求修复而导出仍返回 success:true。表名无法像单元格文本那样静默
  // 清洗（改名会改变业务语义），直接拒绝。
  const invalid = firstInvalidXmlCode(name);
  if (invalid !== null) {
    throw new Error(
      `[excel-exporter] sheet name ${JSON.stringify(
        name,
      )} contains ${formatCodePoint(invalid)}, which XML/OOXML does not allow`,
    );
  }
}

// Excel 工作表硬上限（ECMA-376 / Excel 规格）：1,048,576 行 × 16,384 列。
const EXCEL_MAX_ROWS = 1_048_576;
const EXCEL_MAX_COLS = 16_384;

/**
 * Validate sheet dimensions against Excel's hard limits: 1,048,576 rows
 * (header block included) and 16,384 columns. Beyond these the stream path
 * would write every `<row>` and zip successfully, producing a workbook Excel
 * opens with a repair prompt while the export reports success -- the failure
 * class validateInput exists to eliminate. Checked identically on the
 * Workbook and stream paths.
 *
 * `leafCount` / `headerRowCount` 必须是最终生效值：`exportExcel` 的前置校验在
 * 序号列注入**之前**运行，须自行加上 indexColumn 占用的一列；构建器收到的是
 * 已注入的 sheet，直接传 flattenColumnTree 的结果即可。
 */
export function validateSheetLimits(
  sheet: SheetConfig,
  leafCount: number,
  headerRowCount: number,
): void {
  if (leafCount > EXCEL_MAX_COLS) {
    throw new Error(
      `[excel-exporter] sheet "${sheet.name}" has ${leafCount.toLocaleString("en-US")} columns, exceeding Excel's ${EXCEL_MAX_COLS.toLocaleString("en-US")}-column limit`,
    );
  }
  const totalRows = headerRowCount + sheet.data.length;
  if (totalRows > EXCEL_MAX_ROWS) {
    throw new Error(
      `[excel-exporter] sheet "${sheet.name}" has ${totalRows.toLocaleString("en-US")} rows (header + data), exceeding Excel's ${EXCEL_MAX_ROWS.toLocaleString("en-US")}-row limit`,
    );
  }
}

/**
 * Validate user-supplied merge ranges against the sheet's data area. A merge
 * with a zero/negative span or an out-of-bounds endpoint would produce a
 * reversed or dangling range that Excel treats as a corrupt workbook, so
 * reject it here -- identically on the Workbook and stream paths --
 * with a message naming the offending merge.
 *
 * `MergeRange` is data-relative (row 0 = first data row): bounds are the leaf
 * column count and the data row count. Overlap is checked across the user
 * ranges themselves; they cannot collide with header merges because they sit
 * entirely below the header block.
 */
export function validateMerges(sheet: SheetConfig, leafCount: number): void {
  const merges = sheet.merges;
  // 非数组真值（JS 调用方传 {} 之类）会绕过下方的 ?.length 放行，随后在
  // 两条构建路径的消费端（workbook 的 merges?.forEach / stream 的
  // (merges ?? []).map）都以裸 TypeError 失败，且降级链会把 stream 全量
  // 构建再白跑一遍——与 columns/data 的 Array.isArray 前置拦截对齐。
  // null/undefined 仍视为未提供（消费端本就安全）。
  if (merges != null && !Array.isArray(merges)) {
    throw new Error(
      `[excel-exporter] sheet "${sheet.name}" merges must be an array of MergeRange`,
    );
  }
  if (!merges?.length) return;

  const atOf = (i: number): string => {
    const m = merges[i];
    return `merge #${i} {row: ${m.row}, col: ${m.col}, rowspan: ${m.rowspan}, colspan: ${m.colspan}}`;
  };

  // 第一遍只做形状/边界校验。重叠检测放到排序后的扫描线里，避免原来
  // 对每条 merge 扫描全部前序项（O(n²)：4 万条不重叠合并仅校验就要 ~1s，
  // 导出端到端 12.8s；合并来自用户输入时也是算法级 DoS 面）。
  interface Pending {
    index: number;
    row: number;
    col: number;
    endRow: number;
    endCol: number;
  }
  const pending: Pending[] = [];
  for (let i = 0; i < merges.length; i++) {
    const m = merges[i];
    if (m === null || typeof m !== "object") {
      throw new Error(
        `[excel-exporter] sheet "${sheet.name}" merge #${i} must be an object`,
      );
    }
    const at = atOf(i);
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
    pending.push({
      index: i,
      row: m.row,
      col: m.col,
      endRow: m.row + m.rowspan,
      endCol: m.col + m.colspan,
    });
  }

  // 按行排序后，只需要和"仍覆盖当前行"的合并比列区间。活动集合用
  // 最小堆按 endRow 到期淘汰；典型（互不重叠）输入下活动集合很小，
  // 最坏情况也会在发现首个重叠时立即抛错。稳定性由 index 兜底，保证
  // 错误消息仍指向原数组下标。
  pending.sort((a, b) => a.row - b.row || a.col - b.col || a.index - b.index);
  const active = new Set<Pending>();
  const heap: Pending[] = [];
  const endRowLess = (a: Pending, b: Pending): boolean => a.endRow < b.endRow;
  const heapPush = (item: Pending): void => {
    heap.push(item);
    let i = heap.length - 1;
    while (i > 0) {
      const parent = (i - 1) >> 1;
      if (!endRowLess(heap[i], heap[parent])) break;
      [heap[i], heap[parent]] = [heap[parent], heap[i]];
      i = parent;
    }
  };
  const heapPop = (): Pending => {
    const top = heap[0];
    const last = heap.pop()!;
    if (heap.length > 0) {
      heap[0] = last;
      let i = 0;
      for (;;) {
        const left = i * 2 + 1;
        const right = left + 1;
        let smallest = i;
        if (left < heap.length && endRowLess(heap[left], heap[smallest]))
          smallest = left;
        if (right < heap.length && endRowLess(heap[right], heap[smallest]))
          smallest = right;
        if (smallest === i) break;
        [heap[i], heap[smallest]] = [heap[smallest], heap[i]];
        i = smallest;
      }
    }
    return top;
  };

  for (const cur of pending) {
    while (heap.length > 0 && heap[0].endRow <= cur.row) {
      active.delete(heapPop());
    }
    for (const other of active) {
      if (cur.col < other.endCol && other.col < cur.endCol) {
        // 消息语义与旧实现一致：报告"后出现（原数组下标更大）的那条"
        // 与"先出现的那条"重叠。
        const offender = cur.index > other.index ? cur : other;
        const overlapped = cur.index > other.index ? other : cur;
        throw new Error(
          `[excel-exporter] sheet "${sheet.name}" ${atOf(offender.index)} overlaps merge #${overlapped.index}; merged ranges must be disjoint`,
        );
      }
    }
    active.add(cur);
    heapPush(cur);
  }
}
