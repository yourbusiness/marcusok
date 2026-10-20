import { beforeEach, describe, expect, it, vi } from "vitest";
import type * as XlsxCore from "@marcusok/xlsx-core";
import type { Workbook } from "@marcusok/xlsx-core";

// 直接驱动 parse.worker 的消息协议（此前该文件无任何测试引用：主线程
// worker-client 在 Node 测试环境无法拉起真实 Worker，parse.test 走的是
// parseOnMainThread 回退路径）。与 excel-exporter 的 export-worker.test 同款
// 手法：stub 掉 self、把引擎入口 initWasm/readBuffer 换成 mock，把 worker
// 模块当作纯消息处理函数调用。mock 边界刻意收窄到引擎两个入口——嗅探
// （sniff）、错误归一（engine-errors）、模型构建（model/csv）保持真实，
// 锁定的是 worker 本体的行为：格式分流、wasm 缓存（loadedWasmKey）、
// domain 标记（客户端据此决定是否回退主线程）与 id 协议。
const { initWasmMock, readBufferMock } = vi.hoisted(() => ({
  initWasmMock: vi.fn(),
  readBufferMock: vi.fn(),
}));

vi.mock("@marcusok/xlsx-core", async (importOriginal) => {
  // 其余导出面（getBuiltinFormat 等）保留真实实现：model.ts 的样式链
  // 构建会用到，mock 掉会把被测的模型产出一起换掉
  const actual = await importOriginal<typeof XlsxCore>();
  return { ...actual, initWasm: initWasmMock, readBuffer: readBufferMock };
});

interface Posted {
  data: Record<string, unknown>;
}

const posted: Posted[] = [];
const fakeSelf = {
  onmessage: null as ((e: { data: unknown }) => void) | null,
  postMessage: (data: Record<string, unknown>) => {
    posted.push({ data });
  },
};

/** 重新导入 worker 模块：模块级 wasmReady/loadedWasmKey 状态随之复位。 */
async function freshWorker(): Promise<(e: { data: unknown }) => void> {
  vi.resetModules();
  vi.stubGlobal("self", fakeSelf);
  await import("../worker/parse.worker");
  if (!fakeSelf.onmessage) throw new Error("worker did not install onmessage");
  return fakeSelf.onmessage;
}

function send(
  onmessage: (e: { data: unknown }) => void,
  req: {
    id: number;
    bytes: Uint8Array;
    password?: string;
    wasmUrl?: string;
  },
): Promise<void> {
  return Promise.resolve(
    onmessage({
      data: {
        id: req.id,
        bytes: req.bytes,
        password: req.password,
        wasmUrl: req.wasmUrl,
      },
    }),
  );
}

function responsesOf(id: number): Record<string, unknown>[] {
  return posted.filter((p) => p.data.id === id).map((p) => p.data);
}

const CSV_BYTES = new TextEncoder().encode("a;b\r\n1;2.5\r\n");
// BEL 控制字符流：可打印率远低于 95%，嗅探判 unknown（不触发 wasm）
const JUNK_BYTES = new Uint8Array(2048).fill(7);
const OLE2_BYTES = new Uint8Array([
  0xd0,
  0xcf,
  0x11,
  0xe0,
  0xa1,
  0xb1,
  0x1a,
  0xe1,
  ...new Uint8Array(56),
]);
// PK 头 + 垃圾：嗅探判 zip（走引擎路径）；buildStylesOverlay 的 unzipSync
// 对其抛错并被吞掉（overlay 为 null 的降级路径），不影响协议断言
const ZIP_BYTES = new Uint8Array([
  0x50,
  0x4b,
  0x03,
  0x04,
  ...new Uint8Array(60).fill(1),
]);

/** readBuffer mock 的最小 Workbook（buildPreviewWorkbook 真实消费）。 */
function minimalWorkbook(): Workbook {
  return {
    sheetNames: ["S"],
    getSheet: () => ({
      name: "S",
      state: "visible",
      rows: [
        {
          index: 1,
          cells: [
            {
              reference: "A1",
              value: "1",
              cellType: "number",
              styleIndex: null,
            },
          ],
        },
      ],
      columns: [],
      mergeCells: [],
      frozenPane: null,
      view: null,
    }),
    workbookViews: [],
    dateSystem: "date1900",
    styles: { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [] },
  } as unknown as Workbook;
}

beforeEach(() => {
  vi.clearAllMocks();
  posted.length = 0;
  initWasmMock.mockResolvedValue(undefined);
  readBufferMock.mockResolvedValue(minimalWorkbook());
});

describe("parse.worker 消息协议", () => {
  it("CSV 字节走真实分流（sniff→csvToWorkbook），不经 wasm 直接回 ok", async () => {
    const onmessage = await freshWorker();
    await send(onmessage, { id: 41, bytes: CSV_BYTES });
    const [resp] = responsesOf(41);
    expect(resp).toMatchObject({ id: 41, ok: true });
    expect(resp.workbook).toMatchObject({
      sheets: [{ name: "CSV", rowCount: 2 }],
    });
    expect(initWasmMock).not.toHaveBeenCalled();
  });

  it("未知格式 → UNSUPPORTED + domain:true（确定性结论，客户端免主线程复跑）", async () => {
    const onmessage = await freshWorker();
    await send(onmessage, { id: 42, bytes: JUNK_BYTES });
    expect(responsesOf(42)).toHaveLength(1);
    expect(responsesOf(42)[0]).toMatchObject({
      ok: false,
      code: "UNSUPPORTED",
      domain: true,
    });
  });

  it("OLE2 走引擎，readBuffer 抛 LEGACY_FORMAT → 归一为 LEGACY_FORMAT + domain:true", async () => {
    const onmessage = await freshWorker();
    readBufferMock.mockRejectedValue(
      Object.assign(new Error("legacy biff8"), { code: "LEGACY_FORMAT" }),
    );
    await send(onmessage, { id: 43, bytes: OLE2_BYTES, wasmUrl: "u" });
    expect(responsesOf(43)[0]).toMatchObject({
      ok: false,
      code: "LEGACY_FORMAT",
      domain: true,
    });
  });

  it("readBuffer 抛 PASSWORD_PROTECTED → 密码码直通 + domain:true", async () => {
    const onmessage = await freshWorker();
    readBufferMock.mockRejectedValue(
      Object.assign(new Error("password"), { code: "PASSWORD_PROTECTED" }),
    );
    await send(onmessage, { id: 44, bytes: ZIP_BYTES, wasmUrl: "u" });
    expect(responsesOf(44)[0]).toMatchObject({
      ok: false,
      code: "PASSWORD_PROTECTED",
      domain: true,
    });
  });

  it("initWasm 失败的无码错误 → UNKNOWN + domain:false（允许回退主线程复跑）", async () => {
    // 引擎 initWasm 对 fetch 失败不包错误码（WASM_INIT_FAILED 只在未初始化
    // 守卫抛出）；worker 侧据此产出环境性失败，客户端回退主线程后由
    // parseOnMainThread 重新定码为 WASM——domain 必须为 false
    const onmessage = await freshWorker();
    initWasmMock.mockRejectedValue(new TypeError("fetch failed"));
    await send(onmessage, { id: 45, bytes: ZIP_BYTES, wasmUrl: "u" });
    expect(responsesOf(45)[0]).toMatchObject({
      ok: false,
      code: "UNKNOWN",
      domain: false,
    });
  });

  it("xlsx 成功路径：readBuffer → buildPreviewWorkbook 真实模型回传", async () => {
    const onmessage = await freshWorker();
    await send(onmessage, { id: 46, bytes: ZIP_BYTES, wasmUrl: "u" });
    const [resp] = responsesOf(46);
    expect(resp).toMatchObject({ id: 46, ok: true });
    // buildPreviewWorkbook 真实跑过：模型归一（rows 稀疏化、activeTab 兜底）
    expect(resp.workbook).toMatchObject({
      sheets: [{ name: "S", rowCount: 1, colCount: 1 }],
      activeSheetIndex: 0,
    });
  });

  it("wasm 缓存：同 wasmUrl 不重复 init，换 URL 重新 init（loadedWasmKey）", async () => {
    const onmessage = await freshWorker();
    await send(onmessage, { id: 51, bytes: ZIP_BYTES, wasmUrl: "a" });
    await send(onmessage, { id: 52, bytes: ZIP_BYTES, wasmUrl: "a" });
    await send(onmessage, { id: 53, bytes: ZIP_BYTES, wasmUrl: "b" });
    expect(initWasmMock).toHaveBeenCalledTimes(2);
    expect(initWasmMock).toHaveBeenNthCalledWith(1, "a");
    expect(initWasmMock).toHaveBeenNthCalledWith(2, "b");
    // 三条请求都成功（缓存跳过 init 不影响解析）
    expect(responsesOf(51)[0].ok).toBe(true);
    expect(responsesOf(52)[0].ok).toBe(true);
    expect(responsesOf(53)[0].ok).toBe(true);
  });

  it("并发消息交错：id 各自恰一条响应，互不串扰", async () => {
    const onmessage = await freshWorker();
    await Promise.all([
      send(onmessage, { id: 61, bytes: CSV_BYTES }),
      send(onmessage, { id: 62, bytes: JUNK_BYTES }),
      send(onmessage, { id: 63, bytes: ZIP_BYTES, wasmUrl: "u" }),
    ]);
    expect(responsesOf(61)).toHaveLength(1);
    expect(responsesOf(62)).toHaveLength(1);
    expect(responsesOf(63)).toHaveLength(1);
    expect(responsesOf(61)[0].ok).toBe(true);
    expect(responsesOf(62)[0]).toMatchObject({ code: "UNSUPPORTED" });
    expect(responsesOf(63)[0].ok).toBe(true);
  });

  it("非 Error 的 throw（字符串）→ UNKNOWN 且原始信息不丢", async () => {
    const onmessage = await freshWorker();
    readBufferMock.mockRejectedValue("boom");
    await send(onmessage, { id: 47, bytes: ZIP_BYTES, wasmUrl: "u" });
    expect(responsesOf(47)[0]).toMatchObject({
      ok: false,
      code: "UNKNOWN",
      message: "boom",
    });
  });
});
