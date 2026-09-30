/**
 * @marcusok/progress-overlay（仓库私有包，不再发布 npm）
 *
 * 通用全屏进度遮罩：excel-exporter `overlay` 选项的底层 UI 层，构建期整体
 * 打进其 dist。外部无法（也不应）单独安装本包——手接线场景的公开取用
 * 路径是 `@marcusok/excel-exporter/overlay` 子路径。零运行时依赖，
 * 纯 DOM + 一次注入的样式。
 *
 * ```ts
 * import { showExportOverlay } from "@marcusok/excel-exporter/overlay";
 *
 * const overlay = showExportOverlay({
 *   text: {
 *     title: "正在导出",
 *     phases: { building: "构建中…", downloading: "下载中…" },
 *   },
 * });
 * try {
 *   await work((p, stage) => {
 *     overlay.setProgress(p);
 *     overlay.setPhase(stage);
 *   });
 * } finally {
 *   overlay.close();
 * }
 * ```
 */
export { showProgressOverlay, nextPaint } from "./overlay";
export type {
  ProgressOverlayHandle,
  ProgressOverlayOptions,
  ProgressOverlayTextOptions,
} from "./overlay";
