import type {
  CellStyle,
  ColumnConfig,
  ExportMode,
  ExportOptions,
  ExportPhase,
  MergeRange,
  SheetConfig,
} from "./types";

/**
 * A deliberately dependency-free table column descriptor.
 *
 * It accepts the common field names used by Element Plus (`prop` / `label`,
 * also the library's own naming since 2.2.0) and Ant Design (`dataIndex` /
 * `title`), plus the library's pre-2.2 legacy names (`key` / `header`).
 * `prop` / `label` take precedence when several naming styles are present.
 */
export interface TableColumnInput {
  key?: string;
  dataIndex?: string;
  prop?: string;
  header?: string | number;
  title?: string | number;
  label?: string | number;
  width?: number;
  style?: CellStyle;
  headerStyle?: CellStyle;
  format?: ColumnConfig["format"];
  /** Grouped columns (Ant Design / Element Plus `children`): become a multi-row header. */
  children?: TableColumnInput[];
}

export interface TableSheetInput {
  columns: TableColumnInput[];
  data: Record<string, unknown>[];
  sheetName?: string;
  freezeRows?: number;
  autoFilter?: boolean;
  merges?: MergeRange[];
}

export interface TableExportOptions extends TableSheetInput {
  filename: string;
  mode?: ExportMode;
  onProgress?: (progress: number) => void;
  onPhase?: (phase: ExportPhase, durationMs: number) => void;
  download?: boolean;
}

function normalizeHeader(
  value: string | number | undefined,
  key: string,
): string {
  if (typeof value === "string" && value.length > 0) return value;
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  throw new Error(
    `[excel-exporter] table column "${key}" has no usable header. Provide label, header, or title.`,
  );
}

/**
 * Recursively map a table column (Ant Design / Element Plus shaped) to a
 * `ColumnConfig`. Columns with `children` become group headers (multi-row
 * header); leaves need a usable prop. Group `width`/`style`/`format` are not
 * meaningful (no data cells), so they are dropped.
 */
function toColumnConfig(
  col: TableColumnInput,
  index: number,
  // 检测环必须在转换阶段做：递归 map 先于 exportExcel 的
  // flattenColumnTree 执行，循环 children 若不在本层拦截，用户拿到的是
  // 栈溢出而非 flattenColumnTree 里那个清晰的环错误。
  visiting: Set<TableColumnInput> = new Set(),
): ColumnConfig {
  if (visiting.has(col)) {
    throw new Error(
      "[excel-exporter] circular children reference in table columns",
    );
  }
  // 新名（prop/label）优先，其次本库旧名（key/header），最后 Ant Design 命名兜底。
  const prop = col.prop ?? col.key ?? col.dataIndex;
  const label = normalizeHeader(
    col.label ?? col.header ?? col.title,
    prop ?? `group-${index}`,
  );
  visiting.add(col);
  const children = col.children?.map((child, i) =>
    toColumnConfig(child, i, visiting),
  );
  visiting.delete(col);

  if (children?.length) {
    return {
      label,
      ...(col.headerStyle !== undefined && { headerStyle: col.headerStyle }),
      children,
    };
  }
  if (!prop || typeof prop !== "string") {
    throw new Error(
      `[excel-exporter] table column #${index} has no usable prop. Provide prop, key, or dataIndex.`,
    );
  }
  return {
    prop,
    label,
    ...(col.width !== undefined && { width: col.width }),
    ...(col.style !== undefined && { style: col.style }),
    ...(col.headerStyle !== undefined && { headerStyle: col.headerStyle }),
    ...(col.format !== undefined && { format: col.format }),
  };
}

/**
 * Convert a common table `columns + data` shape into the library's generic
 * `SheetConfig`. This keeps the adapter explicit and easy to test.
 */
export function tableToSheet(input: TableSheetInput): SheetConfig {
  const columns = input.columns.map((col, index) => toColumnConfig(col, index));

  return {
    name: input.sheetName ?? "Sheet1",
    columns,
    data: input.data,
    ...(input.freezeRows !== undefined && { freezeRows: input.freezeRows }),
    ...(input.autoFilter !== undefined && { autoFilter: input.autoFilter }),
    ...(input.merges !== undefined && { merges: input.merges }),
  };
}

/**
 * Convert `exportTable()` options into the generic `ExportOptions` consumed by
 * `exportExcel()`. Kept free of `exportExcel` imports to avoid a circular
 * dependency between the main entrypoint and this adapter.
 */
export function tableExportToOptions(input: TableExportOptions): ExportOptions {
  const {
    columns,
    data,
    sheetName,
    freezeRows,
    autoFilter,
    merges,
    filename,
    mode,
    onProgress,
    onPhase,
    download,
  } = input;

  return {
    filename,
    sheets: [
      tableToSheet({
        columns,
        data,
        sheetName,
        freezeRows,
        autoFilter,
        merges,
      }),
    ],
    ...(mode !== undefined && { mode }),
    ...(onProgress !== undefined && { onProgress }),
    ...(onPhase !== undefined && { onPhase }),
    ...(download !== undefined && { download }),
  };
}
