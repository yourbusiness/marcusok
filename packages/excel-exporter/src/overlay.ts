/**
 * `@marcusok/excel-exporter/overlay` 子路径：历史兼容层。
 *
 * 自 2.8.0 起遮罩已并入主入口（`exportExcel` 默认开启，`overlay: false`
 * 关闭，`overlay: {...}` 定制），能力本体抽到独立包
 * `@marcusok/progress-overlay`（导出/预览等业务包共用的 UI 层）。
 *
 * 本文件只保留两件事：
 * - {@link exportExcelWithOverlay}：旧签名的薄封装，等价于
 *   `exportExcel({ ...options, overlay })`；新代码直接用主入口的 `overlay`
 *   选项即可。
 * - {@link showExportOverlay}：手接线场景（自定义流程自己驱动遮罩）转发到
 *   公共包。注意公共包的进度协议是 `setProgress` / `setPhase(key)`，
 *   文案结构是 `text.phases`（阶段 key → 文案表），旧的
 *   `handlePhase(phase)` 直连与 `text.building` 平铺字段不再提供。
 */

import type { ExportOptions, ExportResult } from "./types";
import { exportExcel } from "./index";
import {
  showProgressOverlay,
  type ProgressOverlayOptions,
} from "@marcusok/progress-overlay";

export type {
  ProgressOverlayHandle as ExportOverlayHandle,
  ProgressOverlayOptions as OverlayOptions,
  ProgressOverlayTextOptions as OverlayTextOptions,
} from "@marcusok/progress-overlay";

/** 手接线场景：转发到公共包（导出语义的默认文案由主入口接线时合并）。 */
export const showExportOverlay = showProgressOverlay;

/**
 * `exportExcel` 的遮罩封装（**兼容保留**，等价于主入口的 `overlay` 选项）：
 * 导出期间显示全局遮罩，结束后（成功或失败）移除。
 *
 * 调用方原有的 `onProgress` / `onPhase` 被**链式追加**而非替换——它们通常还
 * 驱动着自有的指标面板。
 *
 * @example
 * ```ts
 * import { exportExcelWithOverlay } from '@marcusok/excel-exporter/overlay';
 *
 * const result = await exportExcelWithOverlay({
 *   filename: 'report',
 *   sheets: [...],
 * });
 * ```
 */
export async function exportExcelWithOverlay(
  options: ExportOptions,
  overlay: ProgressOverlayOptions = {},
): Promise<ExportResult> {
  // 显式传入的 overlay 定制优先于 options 里可能存在的 overlay 字段。
  return exportExcel({ ...options, overlay });
}
