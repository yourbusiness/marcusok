import { describe, expect, it } from "vitest";
import { toBytes, toLoadBuffer } from "../source";
import { buildPptx, simplePptx } from "./fixtures";

describe("toBytes 防御性复制", () => {
  it("Uint8Array 按视图精确复制（offset/length 视图不拷贝整个底层 buffer）", async () => {
    const base = new Uint8Array(100);
    base.fill(0xff);
    const view = base.subarray(10, 30); // 20 字节视图
    const copy = await toBytes(view);
    expect(copy.length).toBe(20);
    expect(copy[0]).toBe(0xff);
    expect(copy).not.toBe(view);
    // 副本与底层 buffer 解耦：改 base 不影响 copy
    base[10] = 0;
    expect(copy[0]).toBe(0xff);
  });

  it("ArrayBuffer slice 复制，源可继续使用", async () => {
    const src = simplePptx().slice().buffer;
    const copy = await toBytes(src);
    expect(copy.byteLength).toBe(src.byteLength);
    expect(copy.buffer).not.toBe(src);
  });

  it("toLoadBuffer 产出从 0 偏移、独占底层的 ArrayBuffer（引擎会 transfer 它）", () => {
    const view = simplePptx();
    const buf = toLoadBuffer(view);
    expect(buf).toBeInstanceOf(ArrayBuffer);
    expect(buf.byteLength).toBe(view.byteLength);
    // 引擎 transfer buf 后，调用方原 view 不受影响
    structuredClone(view, { transfer: [buf] });
    expect(view.byteLength).toBeGreaterThan(0);
  });

  it("File/Blob 每次读取都是新 buffer", async () => {
    const bytes = simplePptx();
    // Blob 构造要求 ArrayBuffer 底层（TS 5.7+ 的 Uint8Array 泛型收紧）
    const blob = new Blob([bytes.slice().buffer]);
    const a = await toBytes(blob);
    const b = await toBytes(blob);
    expect(a.byteLength).toBe(bytes.byteLength);
    expect(a).not.toBe(b);
  });

  it("同一 Uint8Array source 连续两次 create 得到等价副本", async () => {
    const src = buildPptx({ texts: ["a", "b"] });
    const a = await toBytes(src);
    const b = await toBytes(src);
    expect(a).toEqual(b);
    expect(a).not.toBe(src);
  });
});
