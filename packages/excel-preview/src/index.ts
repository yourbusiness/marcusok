/**
 * @marcusok/excel-preview —— 浏览器端 xlsx 只读预览。
 *
 * 高层 API：`createPreview(container, { source, ... })` 一行接入，内部完成
 * Worker 解析（大文件不冻结 UI）、错误归一与 DOM 虚拟滚动渲染。
 * 低层 API：`parseWorkbookBytes(bytes, { password })` 只出纯数据模型
 * （PreviewWorkbook），供框架封装 / SSR / 自渲染消费。
 */
import type { PreviewInstance, PreviewOptions, PreviewWorkbook } from "./types";
import { SheetRenderer } from "./render/renderer";
import { parseWorkbookSource } from "./worker/worker-client";

export type {
  PreviewCell,
  PreviewCellType,
  PreviewColSpan,
  PreviewFill,
  PreviewFont,
  PreviewBorder,
  PreviewBorderSide,
  PreviewMerge,
  PreviewOptions,
  PreviewParsedInfo,
  PreviewRow,
  PreviewSheet,
  PreviewSource,
  PreviewStyles,
  PreviewWorkbook,
  PreviewXf,
  PreviewInstance,
  PreviewError,
  PreviewErrorCode,
} from "./types";
// WASM/worker 自托管配置面与 excel-exporter 同构（共享 core 的同一 loader）
export { configureWasm, getWasmLoader } from "@marcusok/xlsx-core";
export type { LoaderOptions, LoadState } from "@marcusok/xlsx-core";
// 自定义渲染器可直接复用的单元格格式化器（渲染器同源实现）
export { formatCellValue } from "./numfmt/format";
export type { FormattedValue } from "./numfmt/format";

async function toBytes(source: PreviewOptions["source"]): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  // File/Blob
  const buf = await source.arrayBuffer();
  return new Uint8Array(buf);
}

/**
 * 低层解析：字节 → PreviewWorkbook（不做任何 DOM）。浏览器内经共享
 * Worker，Node/SSR 主线程；CSV 文本自动分流。
 */
export async function parseWorkbookBytes(
  bytes: Uint8Array,
  opts: { password?: string } = {},
): Promise<PreviewWorkbook> {
  const result = await parseWorkbookSource({ bytes, password: opts.password });
  if (!result.ok) {
    const err = new Error(result.message) as Error & { code: string };
    err.code = result.code;
    throw err;
  }
  return result.workbook;
}

/**
 * 挂载预览。
 *
 * @example
 * ```ts
 * import { createPreview } from "@marcusok/excel-preview";
 *
 * const preview = createPreview(el, {
 *   source: file,                 // File / Blob / Uint8Array / ArrayBuffer
 *   onParsed: (info) => console.log(info.sheetNames),
 *   onError: (e) => console.error(e.code, e.message),
 * });
 * preview.setSheet("Sheet2");
 * preview.destroy();
 * ```
 */
export function createPreview(
  container: HTMLElement,
  options: PreviewOptions,
): PreviewInstance {
  const { onParsed, onError } = options;
  let destroyed = false;
  let renderer: SheetRenderer | null = null;
  let workbook: PreviewWorkbook | null = null;

  const boot = async () => {
    const t0 = performance.now();
    try {
      const bytes = await toBytes(options.source);
      const result = await parseWorkbookSource({
        bytes,
        password: options.password,
      });
      if (destroyed) return;
      if (!result.ok) {
        onError?.({ code: result.code, message: result.message });
        return;
      }
      workbook = result.workbook;
      renderer = new SheetRenderer(container);
      const initialSheet =
        typeof options.sheet === "number"
          ? options.sheet
          : options.sheet !== undefined &&
              result.workbook.sheets.some((s) => s.name === options.sheet)
            ? options.sheet
            : undefined;
      renderer.render(result.workbook, {
        sheet: initialSheet,
        showHeaders: options.showHeaders,
        showGridLines: options.showGridLines,
        showTabs: options.showTabs,
        onRendered: (ms) => {
          const sheet = result.workbook.sheets[renderer!.activeSheetIndex];
          onParsed?.({
            sheetNames: result.workbook.sheets.map((s) => s.name),
            sheetCount: result.workbook.sheets.length,
            rowCount: sheet?.rowCount ?? 0,
            colCount: sheet?.colCount ?? 0,
            duration: {
              parse: Math.round(result.parseMs),
              render: Math.round(ms),
              total: Math.round(performance.now() - t0),
            },
          });
        },
      });
    } catch (err) {
      if (destroyed) return;
      onError?.({
        code: "UNKNOWN",
        message: err instanceof Error ? err.message : String(err),
        cause: err,
      });
    }
  };
  void boot();

  return {
    destroy() {
      destroyed = true;
      renderer?.destroy();
      renderer = null;
      workbook = null;
    },
    setSheet(nameOrIndex) {
      if (!renderer) return;
      const idx = renderer.setSheet(nameOrIndex);
      if (idx < 0) {
        onError?.({
          code: "UNKNOWN",
          message: `Sheet not found: ${String(nameOrIndex)}`,
        });
      }
    },
    getSheetNames() {
      return workbook ? workbook.sheets.map((s) => s.name) : [];
    },
  };
}
