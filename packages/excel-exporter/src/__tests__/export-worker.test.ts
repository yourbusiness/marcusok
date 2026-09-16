import { beforeEach, describe, expect, it, vi } from "vitest";

// 直接驱动 export.worker 的消息协议（此前该链路只测过"超时"一支：主线程
// exportInWorker 在 Node 测试环境无法拉起真实 Worker，routing.test 又整体
// mock 掉了边界）。这里 stub 掉 self 与两个引擎依赖，把 worker 模块当作
// 纯消息处理函数调用，锁定两端 WorkerResponse 协议：init/build 阶段消息、
// 终态 ok 响应（含 transfer）、init 幂等（loadedWasmKey）、错误转字符串。
const { initWasmMock, exportAsStreamMock, builderCreateMock } = vi.hoisted(
  () => {
    const exportAsStreamMock = vi.fn();
    const builderCreateMock = vi.fn();
    return {
      initWasmMock: vi.fn(),
      exportAsStreamMock,
      builderCreateMock,
    };
  },
);

vi.mock("modern-xlsx", () => ({ initWasm: initWasmMock }));
vi.mock("../streaming-builder", () => ({
  exportAsStream: exportAsStreamMock,
}));
vi.mock("../workbook-builder", () => ({
  WorkbookBuilder: { create: builderCreateMock },
}));

import type { ExportOptions } from "../types";

interface Posted {
  data: Record<string, unknown>;
  transfer?: Transferable[];
}

const posted: Posted[] = [];
const fakeSelf = {
  onmessage: null as ((e: { data: unknown }) => void) | null,
  postMessage: (data: Record<string, unknown>, transfer?: Transferable[]) => {
    posted.push({ data, transfer });
  },
};

/** 重新导入 worker 模块：模块级 wasmReady/loadedWasmKey 状态随之复位。 */
async function freshWorker(): Promise<(e: { data: unknown }) => void> {
  vi.resetModules();
  vi.stubGlobal("self", fakeSelf);
  await import("../workers/export.worker");
  if (!fakeSelf.onmessage) throw new Error("worker did not install onmessage");
  return fakeSelf.onmessage;
}

const OPTIONS: ExportOptions = {
  filename: "worker-protocol",
  sheets: [
    {
      name: "S",
      columns: [{ key: "a", header: "A" }],
      data: [{ a: 1 }, { a: 2 }],
    },
  ],
};

function send(
  onmessage: (e: { data: unknown }) => void,
  req: {
    id: number;
    mode: "workbook" | "stream";
    wasmUrl?: string;
  },
): Promise<void> {
  onmessage({
    data: {
      id: req.id,
      options: OPTIONS,
      wasmUrl: req.wasmUrl,
      mode: req.mode,
    },
  });
  // worker 的 onmessage 是 async：终态响应 postMessage 之后 promise 才落定。
  return new Promise((resolve) => setTimeout(resolve, 0));
}

beforeEach(() => {
  initWasmMock.mockReset().mockResolvedValue(undefined);
  exportAsStreamMock.mockReset();
  builderCreateMock.mockReset();
  posted.length = 0;
  fakeSelf.onmessage = null;
});

describe("export.worker message protocol", () => {
  it("workbook mode: init once, report init+build phases, ok response transfers bytes", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const toBuffer = vi.fn().mockResolvedValue(bytes);
    builderCreateMock.mockResolvedValue({
      addSheet: vi.fn(),
      toBuffer,
    });

    const onmessage = await freshWorker();
    await send(onmessage, { id: 7, mode: "workbook", wasmUrl: "a.wasm" });

    expect(initWasmMock).toHaveBeenCalledTimes(1);
    expect(initWasmMock).toHaveBeenCalledWith("a.wasm");

    const init = posted[0].data;
    expect(init).toMatchObject({ id: 7, phase: "init" });
    expect(typeof init.duration).toBe("number");

    const build = posted[1].data;
    expect(build).toMatchObject({ id: 7, phase: "build" });
    expect(typeof build.duration).toBe("number");

    const final = posted[2];
    expect(final.data).toEqual({
      id: 7,
      ok: true,
      bytes,
      rowCount: 2,
      engine: "modern-xlsx",
    });
    // bytes.buffer 必须 transfer，不得被结构化克隆复制。
    expect(final.transfer).toEqual([bytes.buffer]);
  });

  it("stream mode: skips wasm init entirely and forwards per-row progress", async () => {
    exportAsStreamMock.mockImplementation(
      (_sheets: unknown, onProgress: (p: number) => void) => {
        onProgress(0.5);
        return Promise.resolve({ bytes: new Uint8Array([9]), rowCount: 42 });
      },
    );

    const onmessage = await freshWorker();
    await send(onmessage, { id: 1, mode: "stream" });

    expect(initWasmMock).not.toHaveBeenCalled();
    const progress = posted[0].data;
    expect(progress).toEqual({ id: 1, progress: 0.5 });

    const final = posted.at(-1)!.data;
    expect(final).toMatchObject({ id: 1, ok: true, rowCount: 42 });
  });

  it("re-inits only when the wasmUrl key changes (string-normalized)", async () => {
    builderCreateMock.mockResolvedValue({
      addSheet: vi.fn(),
      toBuffer: vi.fn().mockResolvedValue(new Uint8Array([0])),
    });

    const onmessage = await freshWorker();
    await send(onmessage, {
      id: 1,
      mode: "workbook",
      wasmUrl: "file:///x/a.wasm",
    });
    await send(onmessage, {
      id: 2,
      mode: "workbook",
      wasmUrl: "file:///x/a.wasm",
    });
    // 同地址的新 URL 对象：worker 端按 String 归一化比较（loadedWasmKey），
    // 不得重复 initWasm 或重复上报 init 阶段。
    await send(onmessage, {
      id: 3,
      mode: "workbook",
      wasmUrl: new URL("file:///x/a.wasm") as unknown as string,
    });
    expect(initWasmMock).toHaveBeenCalledTimes(1);

    await send(onmessage, {
      id: 4,
      mode: "workbook",
      wasmUrl: "file:///x/b.wasm",
    });
    expect(initWasmMock).toHaveBeenCalledTimes(2);
    // id=4 的 init 阶段消息存在（真初始化才上报）。
    expect(posted.some((p) => p.data.id === 4 && p.data.phase === "init")).toBe(
      true,
    );
  });

  it("reports failures as error strings, including non-Error throws", async () => {
    // Error 实例：取 message。
    builderCreateMock.mockRejectedValueOnce(new Error("boom"));
    const onmessage = await freshWorker();
    await send(onmessage, { id: 1, mode: "workbook", wasmUrl: "a.wasm" });
    expect(posted.at(-1)!.data).toEqual({ id: 1, ok: false, error: "boom" });

    // 非 Error 的 throw：String() 兜底，不得变成 undefined（主线程会落到
    // "worker unknown error"，丢失原始信息）。
    builderCreateMock.mockRejectedValueOnce("raw string failure");
    await send(onmessage, { id: 2, mode: "workbook", wasmUrl: "a.wasm" });
    expect(posted.at(-1)!.data).toEqual({
      id: 2,
      ok: false,
      error: "raw string failure",
    });
  });
});
