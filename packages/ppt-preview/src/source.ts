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
 * 都吃 Uint8Array；交给浏览器入口 PptxPresentation.load 前再取
 * `bytes.slice().buffer`（slice 保证从 0 偏移、独占底层 buffer——直接
 * 用视图的 .buffer 可能带偏移且与调用方共享）。
 */
import type { PptSource } from "./types";

/** 归一 source 为独立字节副本（不 detach、不复用调用方 buffer）。 */
export async function toBytes(source: PptSource): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source.slice();
  if (source instanceof ArrayBuffer) return new Uint8Array(source.slice());
  const buf = await source.arrayBuffer();
  return new Uint8Array(buf);
}

/** 从独立副本构造引擎 load 需要的、从 0 偏移且独占底层的 ArrayBuffer。 */
export function toLoadBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.slice().buffer;
}
