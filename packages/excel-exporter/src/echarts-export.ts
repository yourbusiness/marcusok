import type {
  ColumnConfig,
  ExportMode,
  ExportOptions,
  ExportPhase,
  SheetConfig,
} from "./types";

export type EChartsScalar = number | string | null;
// 数组形式即散点坐标（含多维 [x, y, ...dims]——第三维起供 symbolSize /
// visualMap 使用，导出时只取前两维，见 buildItemSheet）。
export type EChartsDatum =
  | EChartsScalar
  | number[]
  | { name?: string; value?: EChartsScalar | number[] };

export interface EChartsSeriesInput {
  name?: string;
  type?: string;
  data?: EChartsDatum[];
}

export interface EChartsXAxisInput {
  type?: string;
  data?: Array<number | string>;
}

/**
 * A minimal structural subset of an ECharts option. The adapter intentionally
 * does not depend on the ECharts runtime or its full type system.
 *
 * `yAxis` shares `EChartsXAxisInput` (an axis is an axis structurally):
 * horizontal bar charts put their categories in `yAxis.data`, so it is
 * consulted as the category source when `xAxis.data` is absent.
 */
export interface EChartsOptionInput {
  xAxis?: EChartsXAxisInput | EChartsXAxisInput[];
  yAxis?: EChartsXAxisInput | EChartsXAxisInput[];
  series?: EChartsSeriesInput[];
  dataset?: unknown;
}

export type EChartsLayout = "wide" | "long";

export interface EChartsSheetInput {
  option: EChartsOptionInput;
  sheetName?: string;
  layout?: EChartsLayout;
  categoryHeader?: string;
  seriesHeader?: string;
  nameHeader?: string;
  valueHeader?: string;
}

export interface EChartsExportOptions extends EChartsSheetInput {
  filename: string;
  mode?: ExportMode;
  onProgress?: (progress: number) => void;
  onPhase?: (phase: ExportPhase, durationMs: number) => void;
  download?: boolean;
}

const DEFAULT_CATEGORY_HEADER = "类目";
const DEFAULT_SERIES_HEADER = "系列";
const DEFAULT_NAME_HEADER = "名称";
const DEFAULT_VALUE_HEADER = "数值";

type ResolvedEChartsSheetInput = EChartsSheetInput & {
  series: EChartsSeriesInput[];
};

function seriesName(series: EChartsSeriesInput, index: number): string {
  return series.name ?? `系列${index + 1}`;
}

/**
 * 长度 ≥2 的全数字数组即散点坐标。ECharts 散点支持多维 `[x, y, ...dims]`
 * （额外维度驱动 symbolSize / visualMap），只认长度恰为 2 会把多维数据
 * 静默挤进 name/value 分支、坐标全部变成 null——静默丢数据。
 */
function isCoordinatePair(value: EChartsDatum): value is number[] {
  return (
    Array.isArray(value) &&
    value.length >= 2 &&
    value.every((item) => typeof item === "number")
  );
}

/**
 * Object-form scatter datum: `{ value: [x, y] }` (an optional `name` is
 * accepted and ignored — the scatter layout has no name column). ECharts
 * accepts this shape interchangeably with the bare `[x, y]` pair, so the
 * coordinate detection below must recognize both or the object form would
 * silently fall through to the name/value branch and stringify the pair.
 */
function isCoordinateDatum(
  value: EChartsDatum,
): value is { name?: string; value: number[] } {
  return (
    value !== null &&
    typeof value === "object" &&
    !Array.isArray(value) &&
    isCoordinatePair(value.value as EChartsDatum)
  );
}

/** A datum of either scatter form: bare pair or object-wrapped pair. */
function isScatterItem(
  value: EChartsDatum,
): value is number[] | { name?: string; value: number[] } {
  return isCoordinatePair(value) || isCoordinateDatum(value);
}

/**
 * Long/item layouts use the header texts themselves as row keys
 * (`{ [seriesHeader]: name, [valueHeader]: item }`), so two identical headers
 * would silently overwrite each other's column. Reject duplicates up front,
 * matching the library's fail-fast input validation.
 */
function assertDistinctHeaders(headers: string[], context: string): void {
  const seen = new Set<string>();
  for (const h of headers) {
    if (seen.has(h)) {
      throw new Error(
        `[excel-exporter] duplicate header "${h}" in ${context}: header texts double as row keys in long/item layouts, so they must be distinct (rename via the *Header options).`,
      );
    }
    seen.add(h);
  }
}

function assertCategoryValue(value: EChartsDatum, context: string): void {
  if (
    isCoordinatePair(value) ||
    (value !== null && typeof value === "object")
  ) {
    throw new Error(
      `[excel-exporter] unsupported ECharts datum for category layout in ${context}. Use long layout for scatter/pie item data.`,
    );
  }
}

function buildCategorySheet(
  input: ResolvedEChartsSheetInput,
  categories: Array<number | string>,
): SheetConfig {
  const series = input.series ?? [];
  const layout = input.layout ?? "wide";
  const categoryHeader = input.categoryHeader ?? DEFAULT_CATEGORY_HEADER;
  const seriesHeader = input.seriesHeader ?? DEFAULT_SERIES_HEADER;
  const valueHeader = input.valueHeader ?? DEFAULT_VALUE_HEADER;

  for (let i = 0; i < series.length; i++) {
    const data = series[i]?.data;
    if (!Array.isArray(data) || data.length !== categories.length) {
      throw new Error(
        `[excel-exporter] ECharts series "${seriesName(series[i], i)}" must have the same length as the category axis data (xAxis.data / yAxis.data).`,
      );
    }
    for (const value of data) assertCategoryValue(value, `series ${i}`);
  }

  if (layout === "long") {
    assertDistinctHeaders(
      [seriesHeader, categoryHeader, valueHeader],
      "category long layout",
    );
    const data: Record<string, unknown>[] = [];
    for (let i = 0; i < series.length; i++) {
      const name = seriesName(series[i], i);
      const values = series[i].data ?? [];
      for (let j = 0; j < categories.length; j++) {
        data.push({
          [seriesHeader]: name,
          [categoryHeader]: categories[j],
          [valueHeader]: values[j],
        });
      }
    }
    return {
      name: input.sheetName ?? "图表数据",
      columns: [
        { prop: seriesHeader, label: seriesHeader },
        { prop: categoryHeader, label: categoryHeader },
        { prop: valueHeader, label: valueHeader },
      ],
      data,
    };
  }

  // Wide layout keys rows by `categoryHeader` plus the internal `__series_N`
  // keys; a user-supplied categoryHeader equal to one of those would be
  // silently overwritten by series data. Reject the collision up front
  // (long/item layouts are already guarded by assertDistinctHeaders).
  if (series.some((_, i) => `__series_${i}` === categoryHeader)) {
    throw new Error(
      `[excel-exporter] categoryHeader "${categoryHeader}" collides with the internal series keys (__series_N) in wide layout; choose a different categoryHeader.`,
    );
  }

  const columns: ColumnConfig[] = [
    { prop: categoryHeader, label: categoryHeader },
    ...series.map((s, i) => ({
      prop: `__series_${i}`,
      label: seriesName(s, i),
    })),
  ];
  const data = categories.map((category, rowIndex) => {
    const row: Record<string, unknown> = { [categoryHeader]: category };
    for (let i = 0; i < series.length; i++) {
      row[`__series_${i}`] = (series[i].data ?? [])[rowIndex];
    }
    return row;
  });

  return {
    name: input.sheetName ?? "图表数据",
    columns,
    data,
  };
}

function buildItemSheet(input: ResolvedEChartsSheetInput): SheetConfig {
  const series = input.series ?? [];
  const seriesHeader = input.seriesHeader ?? DEFAULT_SERIES_HEADER;
  const nameHeader = input.nameHeader ?? DEFAULT_NAME_HEADER;
  const valueHeader = input.valueHeader ?? DEFAULT_VALUE_HEADER;

  const allCoordinate = series.every((s) =>
    (s.data ?? []).every((item) => isScatterItem(item)),
  );
  const anyCoordinate = series.some((s) =>
    (s.data ?? []).some((item) => isScatterItem(item)),
  );

  if (anyCoordinate && !allCoordinate) {
    throw new Error(
      "[excel-exporter] mixing scatter coordinate data with name/value data is not supported by echartsToSheet.",
    );
  }

  if (allCoordinate) {
    const data: Record<string, unknown>[] = [];
    const xKey = "X";
    const yKey = "Y";
    assertDistinctHeaders([seriesHeader, xKey, yKey], "scatter layout");
    // 多维散点的额外维度（第三维起，供 symbolSize / visualMap 使用）进不了
    // 二维表格：静默丢弃会重演"丢数据而 success:true"，导出结束统一告警一次
    // （不是每个数据点都刷一条）。
    let extraDims = false;
    for (let i = 0; i < series.length; i++) {
      const name = seriesName(series[i], i);
      for (const item of series[i].data ?? []) {
        if (!isScatterItem(item)) continue;
        const dims = Array.isArray(item) ? item : item.value;
        if (dims.length > 2) extraDims = true;
        // ECharts 直角坐标系下散点的前两维就是 X/Y。
        data.push({ [seriesHeader]: name, [xKey]: dims[0], [yKey]: dims[1] });
      }
    }
    if (extraDims) {
      console.warn(
        "[excel-exporter] scatter data has dimensions beyond [x, y]; only the X/Y coordinates are exported (extra dims drive symbolSize/visualMap and have no table column).",
      );
    }
    return {
      name: input.sheetName ?? "图表数据",
      columns: [
        { prop: seriesHeader, label: seriesHeader },
        { prop: xKey, label: "X" },
        { prop: yKey, label: "Y" },
      ],
      data,
    };
  }

  assertDistinctHeaders(
    [seriesHeader, nameHeader, valueHeader],
    "name/value layout",
  );
  const data: Record<string, unknown>[] = [];
  for (let i = 0; i < series.length; i++) {
    const name = seriesName(series[i], i);
    for (let j = 0; j < (series[i].data ?? []).length; j++) {
      const item = (series[i].data ?? [])[j];
      if (isScatterItem(item)) continue;
      if (item !== null && typeof item === "object") {
        data.push({
          [seriesHeader]: name,
          [nameHeader]: item.name ?? String(j + 1),
          [valueHeader]: item.value ?? null,
        });
      } else {
        data.push({
          [seriesHeader]: name,
          [nameHeader]: String(j + 1),
          [valueHeader]: item,
        });
      }
    }
  }

  return {
    name: input.sheetName ?? "图表数据",
    columns: [
      { prop: seriesHeader, label: seriesHeader },
      { prop: nameHeader, label: nameHeader },
      { prop: valueHeader, label: valueHeader },
    ],
    data,
  };
}

/**
 * Convert a small structural ECharts option into a `SheetConfig`.
 *
 * Supported shapes:
 * - `xAxis.data`（或水平条形图的 `yAxis.data`）+ multiple one-dimensional
 *   `series[].data` (wide/long).
 * - pie-like `series[].data: { name, value }[]` (long).
 * - scatter-like `series[].data: [x, y][]`、`[x, y, ...dims][]` 或
 *   `{ value: [x, y] }[]` (long) — the ECharts spellings of the same shape,
 *   accepted individually or mixed with each other, but not with name/value
 *   data. 多维散点只导出前两维（X/Y），额外维度以一次 console.warn 告知。
 *
 * `dataset` mode、multiple x/y axes、mixed coordinate/name-value series are
 * rejected explicitly rather than silently producing a misleading table.
 */
export function echartsToSheet(input: EChartsSheetInput): SheetConfig {
  if (input.option.dataset !== undefined) {
    throw new Error(
      "[excel-exporter] ECharts dataset mode is not supported by echartsToSheet. Flatten the data before calling it.",
    );
  }

  const series = input.option.series ?? [];
  if (series.length === 0) {
    throw new Error("[excel-exporter] ECharts option has no series to export.");
  }

  // 双轴图（数组多于一根轴）没有单一的列表格形状：静默取第一根轴会让另一根
  // 轴的 series 以长度不匹配的错误报出来，误导排障——显式拒绝。
  const firstAxis = (
    axis: EChartsXAxisInput | EChartsXAxisInput[] | undefined,
    axisName: "x" | "y",
  ): EChartsXAxisInput | undefined => {
    if (!Array.isArray(axis)) return axis;
    if (axis.length > 1) {
      throw new Error(
        `[excel-exporter] multiple ${axisName} axes are not supported by echartsToSheet (a dual-axis chart has no single-columnar table shape).`,
      );
    }
    return axis[0];
  };
  const xAxis = firstAxis(input.option.xAxis, "x");
  const yAxis = firstAxis(input.option.yAxis, "y");
  // 水平条形图把类目放在 yAxis.data（xAxis 是数值轴、不带 data）。
  const categories = xAxis?.data ?? yAxis?.data;

  if (categories && categories.length > 0) {
    return buildCategorySheet({ ...input, series }, categories);
  }

  return buildItemSheet({ ...input, series });
}

export function echartsExportToOptions(
  input: EChartsExportOptions,
): ExportOptions {
  const {
    option,
    sheetName,
    layout,
    categoryHeader,
    seriesHeader,
    nameHeader,
    valueHeader,
    filename,
    mode,
    onProgress,
    onPhase,
    download,
  } = input;

  return {
    filename,
    sheets: [
      echartsToSheet({
        option,
        sheetName,
        layout,
        categoryHeader,
        seriesHeader,
        nameHeader,
        valueHeader,
      }),
    ],
    ...(mode !== undefined && { mode }),
    ...(onProgress !== undefined && { onProgress }),
    ...(onPhase !== undefined && { onPhase }),
    ...(download !== undefined && { download }),
  };
}
