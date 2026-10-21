/**
 * 导出入口与进度遮罩之间的接线（文案映射 / 下载隔离）。
 *
 * 从 index.ts 抽出，避免主入口继续膨胀；遮罩能力本体在
 * @marcusok/progress-overlay，本模块只承载"导出"这一业务语义。
 */
import type { ExportOptions, ExportPhase } from "./types";
import { triggerDownload } from "./download";
import type { ProgressOverlayTextOptions } from "@marcusok/progress-overlay";

/**
 * Fire-and-forget browser download. The export itself has already succeeded
 * at this point, so a trigger failure (e.g. a sandboxed DOM throwing on
 * a.click()) must not push the caller into the degradation chain and rebuild
 * the whole workbook — the Blob is already in the result. The "download"
 * phase is reported either way.
 */
export function triggerDownloadIsolated(
  options: ExportOptions,
  blob: Blob,
): void {
  const downloadStart = performance.now();
  try {
    triggerDownload(blob, options.filename);
  } catch (e) {
    console.warn(
      `[excel-exporter] download trigger failed; the Blob is still returned in the result. Reason: ${(e as Error).message}`,
    );
  } finally {
    options.onPhase?.("download", performance.now() - downloadStart);
  }
}

/**
 * 遮罩的导出语义文案（默认值）与 ExportPhase → 阶段 key 的映射。
 *
 * onPhase 只在阶段**结束后**回调（见 types.ts），没有"阶段开始"事件，所以
 * 文案只能乐观推进：收到 `init` 即表示 init 已完、当前在做 build。
 */
const OVERLAY_TEXT: ProgressOverlayTextOptions = {
  title: "正在导出 Excel",
  initial: "准备中…",
  phases: {
    building: "正在构建工作簿…",
    downloading: "正在下载…",
    finishing: "即将完成…",
  },
  hint: "数据量较大时可能需要数十秒，请勿关闭页面",
};

export function overlayPhaseKey(
  phase: ExportPhase,
  willDownload: boolean,
): string {
  if (phase === "init") return "building";
  if (phase === "build") return willDownload ? "downloading" : "finishing";
  return "finishing";
}

/** 调用方的 text 片段与导出默认合并：phases 嵌套表单独并，避免整体替换。 */
export function mergeOverlayText(
  caller: ProgressOverlayTextOptions | undefined,
): ProgressOverlayTextOptions {
  return {
    ...OVERLAY_TEXT,
    ...caller,
    phases: { ...OVERLAY_TEXT.phases, ...caller?.phases },
  };
}
