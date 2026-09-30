import {
  Workbook,
  sheetAddAoa,
  encodeCellRef,
  type Worksheet,
  getWasmLoader,
} from "@marcusok/xlsx-core";
import type { CellStyle, SheetConfig, ColumnConfig } from "./types";
import { buildStyleIndex, mergeStyles, BaseCellStyle } from "./style-utils";
import { INDEX_PROP, indexColumnStart } from "./sheet-normalize";
import { flattenColumnTree, columnProp, type HeaderCell } from "./column-tree";
import {
  resolveCellFormat,
  numFormatForSpec,
  validateSheetName,
  validateMerges,
  toStr,
} from "./format-utils";
import { toBlobPart } from "./download";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/**
 * Workbook builder -- batch write path. All data goes through `sheetAddAoa`
 * (array of arrays). For <=50k rows this is the fast, fully-styled path;
 * `Workbook.toBuffer()` is well-behaved here (verified: 50k rows ~700-830ms).
 */
export class WorkbookBuilder {
  private wb: Workbook;
  /**
   * structural CellStyle -> styleIndex. modern-xlsx's StyleBuilder.build
   * appends a fresh font/fill/xf record on every call (no dedup, verified in
   * 1.2.0 dist), so the same style used by N header cells or columns
   * previously produced N identical records. Keyed by JSON.stringify: two
   * structurally-identical styles share an index; a cache miss is harmless.
   */
  private styleIndexCache = new Map<string, number>();

  private constructor() {
    this.wb = new Workbook();
  }

  static async create(): Promise<WorkbookBuilder> {
    await getWasmLoader().ensureLoaded();
    return new WorkbookBuilder();
  }

  addSheet(config: SheetConfig): this {
    const { leaves, headerGrid, headerCells, headerMerges, headerRowCount } =
      flattenColumnTree(config.columns);
    // Same validation as the stream path: invalid merge input must
    // fail with a clear error, not zip a corrupt workbook.
    validateMerges(config, leaves.length);
    // 表名校验前移到数据映射之前：非法表名此前要白付一次 O(rows×cols)
    // 的行映射才报错（校验本身与数据无关）。
    validateSheetName(config.name);

    // Auto-inject an Excel numFormat for typed FormatSpecs (date/datetime/number)
    // so the cell renders correctly without forcing the caller to also set
    // style.numFormat (otherwise dates show as raw serials, numbers as text).
    const columns = leaves.map(withAutoNumFormat);
    // 序号列（INDEX_PROP）的值由行号生成、不读 data——map 的 rowIndex 正好
    // 提供，零数据复制（大数据量下避免 data.map 展开整表对象）。
    const start = indexColumnStart(config);
    const rows = config.data.map((item, rowIndex) =>
      columns.map((col) => {
        // INDEX_PROP 经 columnProp 判定：legacy `key` 别名同样保留（见
        // sheet-normalize 的保留字段说明）
        if (columnProp(col) === INDEX_PROP) return rowIndex + start;
        const v = resolveCellFormat(col, item);
        // Normalize exactly like displayValue on the stream path, so a
        // dataset crossing the 50k threshold (or degrading) keeps identical cell
        // content: non-finite numbers (not valid xsd:double; <v>NaN</v> corrupts
        // the workbook), objects (modern-xlsx would String() them into
        // "[object Object]" instead of JSON) and Dates (localized long form
        // instead of ISO) all become the same visible strings on every path.
        if (typeof v === "number") return Number.isFinite(v) ? v : toStr(v);
        if (typeof v === "boolean") return v;
        // 归一后的空串 = 缺失值（format-utils 的 isBlankString 口径），写出
        // 物理缺失格而非空文本格：引擎 writeAoaRow 对 null 跳格（不建 <c>），
        // 与 stream 路径 fast-xlsx 的 v === "" continue 语义对齐。空文本格
        // ISBLANK()=FALSE，同一份数据跨 50k 阈值会让 COUNTA/ISBLANK 类公式
        // 结果漂移，违反 types.ts 的跨路径同一性契约。
        if (typeof v === "string") return v === "" ? null : v;
        return toStr(v);
      }),
    );
    const aoa = [...headerGrid, ...rows];

    const ws = this.wb.addSheet(config.name);
    sheetAddAoa(ws, aoa, { origin: "A1" });

    return this.applyLayout(
      ws,
      config,
      columns,
      headerCells,
      headerMerges,
      headerRowCount,
      rows.length,
    );
  }

  private applyLayout(
    ws: Worksheet,
    config: SheetConfig,
    columns: ColumnConfig[], // flattened leaf columns (numFormat-injected)
    headerCells: HeaderCell[], // every header cell, for header styling
    headerMerges: HeaderCell[], // span > 1 header cells, for <mergeCell>
    headerRowCount: number,
    dataRowCount: number,
  ): this {
    // Column widths (1-based) -- leaf columns only. The injected index column
    // carries width only when the user set one (see applyIndexColumn): its
    // documented default of 6 is applied here, at the only place that consumes
    // widths -- the stream path never outputs column widths and its
    // feature-detect must not see the library's own default as user config.
    columns.forEach((c, i) => {
      const width = columnProp(c) === INDEX_PROP ? (c.width ?? 6) : c.width;
      if (width !== undefined) ws.setColumnWidth(i + 1, width);
    });

    // Header styles. Column-level headerStyle wins over the sheet-level default.
    // The top-left cell of every header cell carries the style; merged group
    // headers inherit it across the merged region (OOXML styles the anchor cell).
    // Note: `ws.rows[r].cells[c]` cannot be used here -- modern-xlsx packs a
    // row's cells densely, so a header row with merge-covered gaps (multi-row
    // headers) misaligns `cells[c]` from absolute column c. Resolve by ref.
    headerCells.forEach((cell) => {
      // BaseCellStyle 铺在最底层：列级 headerStyle 整体替换表级默认的语义
      // 不变（两者之间不合并），只是两者都没声明对齐时由基底补上居中。
      const headerStyle = mergeStyles(
        BaseCellStyle,
        cell.column.headerStyle ?? config.headerStyle,
      )!;
      const idx = this.cachedStyleIndex(headerStyle);
      const target = ws.cell(encodeCellRef(cell.row, cell.col));
      if (target) target.styleIndex = idx;
    });

    // Column styles: apply to data cells only, matching the `style: not the
    // label` contract in types.ts. Header styling is handled separately above
    // via headerStyle. Sheet-level dataStyle is the base layer, the column's
    // own style deep-merges over it (see mergeStyles), and BaseCellStyle sits
    // underneath both. Data rows start at sheet row headerRowCount (0-based),
    // so slice(headerRowCount) iterates only data rows; mutating styleIndex is
    // a plain JS property write, bypassing ws.cell(ref) ref-parsing overhead.
    // 列样式先按列算好（经 cachedStyleIndex 去重），再按行主序遍历真实存在
    // 的单元格、按 cell.reference 解出的真实列号取样式。不能按
    // `row.cells[i]` 位置索引取格：writeAoaRow 对缺失值（null，含归一后的
    // 空串）跳格，row.cells 是稠密打包数组，cells[i] 不对应第 i 列——行内
    // 任一前列缺失即整体左移，样式会静默写到错误的列上（与上方表头路径
    // 注释指出的同一引擎行为；toStr 的 toJSON 兜底注释也记录过同款错位）。
    const dataRows = ws.rows.slice(headerRowCount);
    const colStyleIdx = columns.map((c) =>
      // 基底恒非空，mergeStyles 在 base 存在时必返回样式——用 ! 收窄
      this.cachedStyleIndex(
        mergeStyles(BaseCellStyle, mergeStyles(config.dataStyle, c.style))!,
      ),
    );
    for (const row of dataRows) {
      for (const cell of row.cells) {
        if (typeof cell.reference !== "string") continue;
        const col = refCol(cell.reference);
        if (col >= 0 && col < colStyleIdx.length) {
          cell.styleIndex = colStyleIdx[col];
        }
      }
    }

    // Freeze header rows
    if (config.freezeRows && config.freezeRows > 0) {
      ws.frozenPane = { rows: config.freezeRows, cols: 0 };
    }

    // Auto-filter over the last header row .. last data row
    if (config.autoFilter) {
      const lastCol = encodeCellRef(0, columns.length - 1).match(/[A-Z]+/)![0];
      ws.autoFilter = `A${headerRowCount}:${lastCol}${headerRowCount + dataRowCount}`;
    }

    // Header merges: rows are already sheet-relative (0-based).
    headerMerges.forEach((m) => {
      const start = encodeCellRef(m.row, m.col);
      const end = encodeCellRef(m.row + m.rowSpan - 1, m.col + m.colSpan - 1);
      ws.addMergeCell(`${start}:${end}`);
    });

    // Data merges: MergeRange is data-relative (row 0 = first data row); add
    // headerRowCount to reach the sheet.
    config.merges?.forEach((m) => {
      const start = encodeCellRef(headerRowCount + m.row, m.col);
      const end = encodeCellRef(
        headerRowCount + m.row + m.rowspan - 1,
        m.col + m.colspan - 1,
      );
      ws.addMergeCell(`${start}:${end}`);
    });

    return this;
  }

  /**
   * buildStyleIndex wrapper that returns a cached index for structurally
   * identical styles (see styleIndexCache) instead of appending duplicates.
   */
  private cachedStyleIndex(style: CellStyle): number {
    const key = JSON.stringify(style);
    const cached = this.styleIndexCache.get(key);
    if (cached !== undefined) return cached;
    const idx = buildStyleIndex(this.wb, style);
    this.styleIndexCache.set(key, idx);
    return idx;
  }

  /** Serialize to Uint8Array (async, avoids sync writeBlob blocking main thread). */
  async toBuffer(): Promise<Uint8Array> {
    return this.wb.toBuffer();
  }

  /** Convenience: serialize and wrap in a Blob. */
  async toBlob(): Promise<Blob> {
    const bytes = await this.toBuffer();
    return new Blob([toBlobPart(bytes)], { type: XLSX_MIME });
  }
}

/**
 * "BC3" → 1（0-based 列号）。只解析前导字母：全表数据格遍历下比
 * decodeCellRef 的完整往返便宜（excel-preview 侧 model.ts 的同款快路径）。
 */
function refCol(reference: string): number {
  let col = 0;
  for (let i = 0; i < reference.length; i++) {
    const c = reference.charCodeAt(i);
    if (c >= 65 && c <= 90) col = col * 26 + (c - 64);
    else if (c >= 97 && c <= 122) col = col * 26 + (c - 96);
    else break;
  }
  return col - 1;
}

/**
 * If a column has a typed FormatSpec (date/datetime/number) but no explicit
 * style.numFormat, inject the matching Excel numFormat so the value displays
 * correctly. Explicit numFormat on the column style always wins.
 *
 * 语义澄清：注入结果落在列级 style，经 mergeStyles 恒压过表级
 * dataStyle.numFormat（override 无条件胜出）。date/datetime 列必须如此——
 * 否则序列号会被表级装饰格式（如 "#,##0"）渲染错；number 列同理不继承
 * 表级千分位，需要千分位时在 format spec 里声明 thousands: true。
 */
function withAutoNumFormat(c: ColumnConfig): ColumnConfig {
  const spec = typeof c.format === "object" ? c.format : null;
  const nf = spec ? numFormatForSpec(spec) : null;
  if (nf && !c.style?.numFormat) {
    return { ...c, style: { ...(c.style ?? {}), numFormat: nf } };
  }
  return c;
}
