/**
 * 解析客户端：浏览器优先共享 Worker（大文件解析不冻结 UI，实测 10 万行×10
 * 列 readBuffer ≈ 1.5s）；Worker 不可用/失败时回退主线程解析（Node/SSR/测试
 * 环境的常态路径）。两条路径产出同一份 PreviewWorkbook 模型。
 */
import { defaultWasmUrl, getWasmLoader, readBuffer } from "@marcusok/xlsx-core";
import { buildPreviewWorkbook } from "../parse/model";
import { csvToWorkbook } from "../parse/csv";
import { normalizeEngineError } from "../parse/engine-errors";
import { sniffFormat } from "../parse/sniff";
import type { PreviewErrorCode, PreviewWorkbook } from "../types";

export interface ParseRequest {
  bytes: Uint8Array;
  password?: string;
}

export interface ParseOk {
  ok: true;
  workbook: PreviewWorkbook;
  parseMs: number;
}
export interface ParseErr {
  ok: false;
  code: PreviewErrorCode;
  message: string;
  cause?: unknown;
  /**
   * true = 对相同字节必然复现的域错误（worker 标记，见 parse.worker 协议）：
   * 解析入口收到后直传调用方，不回退主线程复跑（复跑必然同错，大文件在
   * 主线程整段重跑正是回退机制要规避的冻结源）。
   */
  domain?: boolean;
}
export type ParseResult = ParseOk | ParseErr;

/** 与 parse.worker 相同的错误归一（共享 parse/engine-errors，主线程回退路径复用）。 */
export function normalizeParseError(err: unknown): ParseErr {
  const info = normalizeEngineError(err);
  return { ok: false, ...info, cause: err };
}

/** 共享 Worker（模块级单例；脚本加载失败时按次回退主线程）。 */
let worker: Worker | null = null;
let workerBroken = false;
let requestIdSeq = 0;
const pending = new Map<number, (r: ParseResult) => void>();

function getOrCreateWorker(): Worker | null {
  if (workerBroken) return null;
  if (worker) return worker;
  if (typeof Worker === "undefined" || typeof window === "undefined") {
    return null;
  }
  // 预览的 worker URL 走独立字段 parseWorkerUrl，不读导出包的 workerUrl：
  // 共享 loader（core 单例）服务同页多个业务包，单字段会让两包互拿对方的
  // worker 脚本（预览会拉起 export worker）。
  const { parseWorkerUrl } = getWasmLoader().getOptions();
  // 提升到变量再构造（与 excel-exporter 同因）：内联 new URL 字面量会让
  // 消费方 bundler 把 worker 当入口重新打包（Vite 5 下直接失败）；提升后
  // 仅匹配资产 URL 模式，产物按原样作为 hash 资产输出。
  const url = parseWorkerUrl ?? new URL("./parse.worker.js", import.meta.url);
  const w = new Worker(url, { type: "module" });
  w.onmessage = (e: MessageEvent) => {
    const data = e.data as {
      id: number;
      ok: boolean;
      workbook?: PreviewWorkbook;
      code?: string;
      message?: string;
      domain?: boolean;
    };
    const resolve = pending.get(data.id);
    if (!resolve) return;
    pending.delete(data.id);
    if (data.ok && data.workbook) {
      resolve({ ok: true, workbook: data.workbook, parseMs: 0 });
    } else {
      resolve({
        ok: false,
        code: (data.code as PreviewErrorCode) ?? "UNKNOWN",
        message: data.message ?? "worker parse failed",
        domain: data.domain === true,
      });
    }
  };
  w.onerror = () => {
    // 脚本加载失败（自托管路径配错等）：终止并回退主线程路径
    w.terminate();
    if (worker === w) {
      worker = null;
      workerBroken = true;
    }
    for (const [id, resolve] of pending) {
      pending.delete(id);
      resolve({
        ok: false,
        code: "UNKNOWN",
        message: "parse worker failed to load; falling back to main thread",
      });
    }
  };
  worker = w;
  return w;
}

/** 主线程解析（Node/SSR/测试/Worker 失败回退共用）。 */
export async function parseOnMainThread(
  req: ParseRequest,
): Promise<ParseResult> {
  const t0 = performance.now();
  try {
    const form = sniffFormat(req.bytes);
    if (form === "text") {
      const workbook = csvToWorkbook(req.bytes);
      return { ok: true, workbook, parseMs: performance.now() - t0 };
    }
    if (form === "unknown") {
      return {
        ok: false,
        code: "UNSUPPORTED",
        message:
          "Unrecognized file format: expected an .xlsx/.xlsm (ZIP) or .csv (text) file.",
      };
    }
    const loader = getWasmLoader();
    if (!loader.supported) {
      return {
        ok: false,
        code: "WASM",
        message: "WebAssembly is not supported in this environment.",
      };
    }
    await loader.ensureLoaded();
    const wb = await readBuffer(
      req.bytes,
      req.password ? { password: req.password } : undefined,
    );
    const workbook = buildPreviewWorkbook(wb, req.bytes);
    return { ok: true, workbook, parseMs: performance.now() - t0 };
  } catch (err) {
    return normalizeParseError(err);
  }
}

/**
 * 解析入口：浏览器走共享 Worker；Node/SSR 等无 Worker 环境主线程。Worker
 * 失败（脚本加载失败/内部错误）时以未转移的原始字节自动重试主线程一次。
 */
export async function parseWorkbookSource(
  req: ParseRequest,
): Promise<ParseResult> {
  const w = getOrCreateWorker();
  if (!w) return parseOnMainThread(req);

  const t0 = performance.now();
  const id = ++requestIdSeq;
  const { wasmUrl } = getWasmLoader().getOptions();
  const result = await new Promise<ParseResult>((resolve) => {
    pending.set(id, resolve);
    // wasm URL 显式转发：worker 内相对解析不可靠（脚本可能被复制/重命名）。
    // bytes 不转移（结构化克隆复制）：几 MB 文件的复制成本毫秒级，换取
    // worker 失败后主线程兜底仍持有原始字节——转移会把 buffer 中性化，
    // 兜底即失效。
    try {
      w.postMessage({
        id,
        bytes: req.bytes,
        password: req.password,
        // 必须 String()：URL 对象不可结构化克隆，postMessage 会同步抛
        // DataCloneError（Chrome/Edge 实测），而默认路径正是 new URL(...)。
        wasmUrl: String(wasmUrl ?? defaultWasmUrl()),
      });
    } catch (err) {
      // 同步失败（不可克隆载荷等）：先清 pending 槽位（否则该 id 永久悬留），
      // 再按 worker 失败上报，交给下方主线程兜底——与 onerror 同一降级语义，
      // 不让一个可规避的载荷问题变成解析失败。
      pending.delete(id);
      resolve({
        ok: false,
        code: "UNKNOWN",
        message:
          "parse worker postMessage failed: " +
          (err instanceof Error ? err.message : String(err)),
      });
    }
  });
  if (result.ok) {
    return { ...result, parseMs: performance.now() - t0 };
  }
  // 域错误（密码/格式判别等确定性结论）直传：主线程复跑必然同错，
  // 回退只会把加密/损坏文件的解析成本翻倍
  if (result.domain) return result;
  // Worker 失败（脚本加载失败/内部错误）→ 主线程兜底（bytes 未转移，可用）
  return parseOnMainThread(req);
}
