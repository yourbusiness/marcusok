/**
 * Fast stream 不支持、会被静默丢弃的配置特性。
 *
 * 单独成模块是为了让"日志警告"与"返回给调用方的结构化信号"使用同一份
 * 判定：正常 auto ≥ 50k 路由和失败降级都会丢样式，调用方需要能程序化
 * 区分，而不是只靠 console.warn。
 */
import type { SheetConfig } from "./types";
import { someColumn } from "./column-tree";

/** 单表的被丢弃特性（顺序即警告文案顺序）。 */
export function skippedFeatures(config: SheetConfig): string[] {
  const skipped: string[] = [];
  // someColumn walks the whole tree: width/style/headerStyle may sit on
  // nested nodes, and a top-level-only scan would drop them silently.
  if (someColumn(config.columns, (c) => c.width !== undefined)) {
    skipped.push("width");
  }
  if (
    config.headerStyle !== undefined ||
    someColumn(config.columns, (c) => c.headerStyle !== undefined)
  ) {
    skipped.push("headerStyle");
  }
  if (someColumn(config.columns, (c) => c.style !== undefined)) {
    skipped.push("style");
  }
  if (config.dataStyle !== undefined) skipped.push("dataStyle");
  if (config.freezeRows) skipped.push("freezeRows");
  if (config.autoFilter) skipped.push("autoFilter");
  return skipped;
}

/** 全部 sheet 的被丢弃特性并集（保持首次出现顺序）。 */
export function streamDroppedFeatures(sheets: SheetConfig[]): string[] {
  const dropped = new Set<string>();
  for (const sheet of sheets) {
    for (const feature of skippedFeatures(sheet)) dropped.add(feature);
  }
  return [...dropped];
}
