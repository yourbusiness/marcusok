/**
 * @marcusok/progress-overlay
 *
 * 通用全屏进度遮罩：@marcusok 各业务包（excel-exporter 及未来的导出/预览
 * 包）共用的 UI 层，也可独立使用。零运行时依赖，纯 DOM + 一次注入的样式。
 *
 * ```ts
 * import { showProgressOverlay } from "@marcusok/progress-overlay";
 *
 * const overlay = showProgressOverlay({
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
