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
  /**
   * true = 环境性失败但禁止回退主线程（worker 超时专用）：worker 已挂死
   * （wasm 初始化悬置/引擎死循环）时，主线程整段重跑同样的大文件正是要
   * 规避的冻结源、且大概率同样挂死。与 domain 分立：超时不是对相同字节
   * 必然复现的结论，只是本次环境性失败，调用方重试会重建新 worker。
   */
  noFallback?: boolean;
}
export type ParseResult = ParseOk | ParseErr;

/** 与 parse.worker 相同的错误归一（共享 parse/engine-errors，主线程回退路径复用）。 */
export function normalizeParseError(err: unknown): ParseErr {
  const info = normalizeEngineError(err);
  return { ok: false, ...info, cause: err };
}

/**
 * 共享 Worker（模块级单例）。脚本加载或构造失败置 workerBroken，此后永久
 * 走主线程——worker URL 是固定配置、失败对后续请求必然复现，重试只在每次
 * 解析前多付一次构造/加载（超时挂死不置位：通常与特定负载相关，下次请求
 * 重建新实例）。
 */
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
  // loader 源码由两个业务包共用（各自打包一份单例），单字段会让两包互拿
  // 对方的 worker 脚本（预览会拉起 export worker）。
  const { parseWorkerUrl } = getWasmLoader().getOptions();
  // 提升到变量再构造（与 excel-exporter 同因）：内联 new URL 字面量会让
  // 消费方 bundler 把 worker 当入口重新打包（Vite 5 下直接失败）；提升后
  // 仅匹配资产 URL 模式，产物按原样作为 hash 资产输出。
  const url = parseWorkerUrl ?? new URL("./parse.worker.js", import.meta.url);
  // 构造守卫：跨域脚本（如把 parseWorkerUrl 配到裸 CDN 域名）在 Chrome 下
  // 同步抛 SecurityError——worker 脚本受同源限制。异常若穿透到调用方，
  // parseWorkbookSource 的主线程回退不会生效（本模块的失败约定是 resolve
  // 错误结果而非 reject），必须就地吞掉、按 onerror 同语义收尾。
  let w: Worker;
  try {
    w = new Worker(url, { type: "module" });
  } catch (err) {
    workerBroken = true;
    console.warn(
      "[excel-preview] parse worker failed to construct: " +
        (err instanceof Error ? err.message : String(err)) +
        ". Falling back to main-thread parsing. Worker scripts must be " +
        "same-origin — serve parseWorkerUrl from your own origin (a raw " +
        "CDN URL throws at construction; only the WASM binary may be cross-origin).",
    );
    return null;
  }
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
    try {
      await loader.ensureLoaded();
    } catch (err) {
      // 加载/初始化失败（资产 404、CSP 禁 WebAssembly、网络、二进制不匹配）：
      // xlsx-core 抛的是不带 code 的普通 Error，直接上抛会被归一成 UNKNOWN、
      // 调用方无法按码分流。这里补上 WASM 码，交统一映射产出友好文案（原始
      // 错误经 cause 链保留，诊断信息不丢）。
      throw Object.assign(
        new Error(err instanceof Error ? err.message : String(err)),
        { code: "WASM" },
      );
    }
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

/** 拒绝全部在途请求（worker 超时/报废时整批收尾，槽位一并清空）。 */
function rejectAllPending(err: Omit<ParseErr, "ok">): void {
  for (const [id, resolve] of pending) {
    pending.delete(id);
    resolve({ ok: false, ...err });
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
  const { wasmUrl, workerTimeoutMs } = getWasmLoader().getOptions();
  let timer: ReturnType<typeof setTimeout> | undefined;
  const result = await new Promise<ParseResult>((resolve) => {
    pending.set(id, resolve);
    // 每请求一个计时器（首个触发的负责收尾）：worker 挂死（wasm 初始化
    // 悬置/引擎死循环）时 pending 会无限等待，onParsed/onError 均不触发。
    // 超时即终止共享 worker 并拒绝全部在途请求——与 xlsx-core 对
    // workerTimeoutMs 的文档语义一致（超时终止共享 worker、兄弟请求一并
    // 拒绝，excel-exporter 的 export worker 同款行为）。超时错误标记
    // noFallback（见 ParseErr）：挂死后主线程重跑同样的大文件正是要规避
    // 的冻结源。worker 引用置空但不置 workerBroken——挂死通常与特定负载
    // 相关，下次请求重建新实例，调用方重试即恢复 worker 路径。
    timer = setTimeout(() => {
      w.terminate();
      if (worker === w) worker = null;
      rejectAllPending({
        code: "UNKNOWN",
        message: `parse worker timed out after ${workerTimeoutMs ?? 120_000}ms (worker terminated); retry the parse`,
        noFallback: true,
      });
    }, workerTimeoutMs ?? 120_000);
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
      // 同步失败（不可克隆载荷等）：先清 pending 槽位与计时器（否则该 id
      // 永久悬留/计时器空转），再按 worker 失败上报，交给下方主线程兜底
      // ——与 onerror 同一降级语义，不让一个可规避的载荷问题变成解析失败。
      clearTimeout(timer);
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
  // 响应已到（含 onerror/超时收尾后 resolve 的情形）：清掉本请求的计时器，
  // 已触发的计时器 clearTimeout 为无害 no-op
  clearTimeout(timer);
  if (result.ok) {
    return { ...result, parseMs: performance.now() - t0 };
  }
  // 域错误（密码/格式判别等确定性结论）直传：主线程复跑必然同错，
  // 回退只会把加密/损坏文件的解析成本翻倍
  if (result.domain) return result;
  // worker 超时：不回退主线程（见 ParseErr.noFallback）
  if (result.noFallback) return result;
  // Worker 失败（脚本加载失败/内部错误）→ 主线程兜底（bytes 未转移，可用）
  return parseOnMainThread(req);
}
