/**
 * source 归一化与防御性复制。
 *
 * 引擎的解析请求会把输入 ArrayBuffer 放进 postMessage 的 transfer list
 * （上游 chunk 内 `transfer:[buffer]`，spike 实证），transfer 后调用方
 * 持有的原始 buffer 立即 detach（byteLength=0）。本包必须在 sniff 与
 * 引擎之间持有独立副本：
 * - Uint8Array —— 按 offset/length 精确复制（slice 自动处理视图偏移，
 *   不复制整个底层 buffer，避免共享 buffer 的越界拷贝）；
 * - ArrayBuffer —— slice 复制（同 buffer 可能被调用方复用/传输）；
 * - File / Blob —— 每次 arrayBuffer() 都是新读，无需再复制。
 *
 * 产出统一为 Uint8Array：sniff 与 Node 解析入口（openPptxPresentation）
 * 都吃 Uint8Array；交给浏览器入口 PptxPresentation.load 前由 toLoadBuffer
 * 取出底层 ArrayBuffer——toBytes 已保证从 0 偏移、独占底层 buffer，可直接
 * transfer，不再复制第二份（仅非全量视图才兜底 slice）。
 */
import type { PptSource } from "./types";

/**
 * 跨 realm 判定：`instanceof` 对 iframe/Node vm 的构造器恒为 false，
 * 会让跨 realm 的 TypedArray 落到 `.arrayBuffer()`（Uint8Array 没有该
 * 方法）而报错。ArrayBuffer.isView 按内部槽判定，跨 realm 有效。
 */
function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === "[object ArrayBuffer]";
}

/** 归一 source 为独立字节副本（不 detach、不复用调用方 buffer）。 */
export async function toBytes(source: PptSource): Promise<Uint8Array> {
  if (ArrayBuffer.isView(source)) {
    return new Uint8Array(
      source.buffer,
      source.byteOffset,
      source.byteLength,
    ).slice();
  }
  if (isArrayBuffer(source)) return new Uint8Array(source).slice();
  const buf = await source.arrayBuffer();
  return new Uint8Array(buf);
}

/** 从独立副本构造引擎 load 需要的、从 0 偏移且独占底层的 ArrayBuffer。
 *
 * toBytes 的三个分支都产出从 0 偏移、独占底层 buffer 的精确长度视图，
 * 因此可以直接把底层 buffer 交给引擎 transfer——再 slice() 一次会让
 * 加载期间的峰值内存翻倍（大文件/移动端显著）。仅在视图非全量时兜底复制。 */
export function toLoadBuffer(bytes: Uint8Array): ArrayBuffer {
  if (bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) {
    return bytes.buffer as ArrayBuffer;
  }
  return bytes.slice().buffer;
}
