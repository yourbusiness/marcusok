/**
 * 解析客户端：浏览器优先共享 Worker（大文件解析不冻结 UI，实测 10 万行×10
 * 列 readBuffer ≈ 1.5s）；Worker 不可用/失败时回退主线程解析（Node/SSR/测试
 * 环境的常态路径）。两条路径产出同一份 PreviewWorkbook 模型。
 *
 * Worker 请求走显式 FIFO 队列：同一时刻只派发一个请求，超时从 postMessage
 * 成功后开始计时（只覆盖 worker 实际处理时间，不含排队等待）。超时只收尾
 * 当前请求并重建 worker，排队的请求继续落到新实例；排队/中止不再连坐。
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
  /**
   * 取消信号：中止即以 code ABORTED 解除调用方等待。语义边界——主线程
   * 路径在 await 边界（ensureLoaded / readBuffer 之后）放弃后续大计算，
   * 同步 WASM 段不可中途打断；worker 路径排队中的请求直接摘除，活动请求
   * 终止共享 worker（排队的其他请求会落到新实例），迟到结果经已清的槽位
   * 丢弃。取消结果不回退主线程复跑。
   */
  signal?: AbortSignal;
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

/** 共享 Worker（模块级单例）。脚本加载或构造失败置 workerBroken，此后永久
 * 走主线程——worker URL 是固定配置、失败对后续请求必然复现，重试只在每次
 * 解析前多付一次构造/加载（超时挂死不置位：通常与特定负载相关，下次请求
 * 重建新实例）。`terminateParseWorker()` 会清除该标志，以便换 URL 后按新
 * 配置重建。 */
let worker: Worker | null = null;
let workerBroken = false;
let requestIdSeq = 0;

interface QueueEntry {
  id: number;
  req: ParseRequest;
  resolve: (r: ParseResult) => void;
  worker: Worker | null;
  timer?: ReturnType<typeof setTimeout>;
  onAbort?: () => void;
}

let active: QueueEntry | null = null;
const queue: QueueEntry[] = [];

/** 调用方主动取消的统一结果（ABORTED，不回退主线程复跑——复跑违背取消意图）。 */
function abortedResult(): ParseErr {
  return {
    ok: false,
    code: "ABORTED",
    message:
      "The parse was cancelled by the caller (createPreview destroy() or an aborted signal).",
  };
}

function workerLoadFailure(): ParseErr {
  return {
    ok: false,
    code: "UNKNOWN",
    // non-domain -> parseWorkbookSource falls back to the main thread.
    message: "parse worker failed to load; falling back to main thread",
  };
}

function finishEntry(entry: QueueEntry): void {
  if (entry.timer !== undefined) {
    clearTimeout(entry.timer);
    entry.timer = undefined;
  }
  if (entry.req.signal && entry.onAbort) {
    entry.req.signal.removeEventListener("abort", entry.onAbort);
    entry.onAbort = undefined;
  }
  if (active === entry) active = null;
}

function resolveEntry(entry: QueueEntry, result: ParseResult): void {
  finishEntry(entry);
  entry.resolve(result);
}

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
  worker = w;
  w.onmessage = (e: MessageEvent) => handleMessage(w, e.data);
  w.onerror = () => handleWorkerError(w);
  return w;
}

function handleMessage(w: Worker, data: unknown): void {
  const msg = data as {
    id: number;
    ok: boolean;
    workbook?: PreviewWorkbook;
    code?: string;
    message?: string;
    domain?: boolean;
  };
  // Responses from a replaced/terminated worker must not settle the request
  // now running on its replacement.
  if (!active || active.worker !== w || active.id !== msg.id) return;
  const entry = active;
  if (msg.ok && msg.workbook) {
    resolveEntry(entry, { ok: true, workbook: msg.workbook, parseMs: 0 });
  } else {
    resolveEntry(entry, {
      ok: false,
      code: (msg.code as PreviewErrorCode) ?? "UNKNOWN",
      message: msg.message ?? "worker parse failed",
      domain: msg.domain === true,
    });
  }
  dispatchNext();
}

function handleWorkerError(w: Worker): void {
  w.terminate();
  if (worker === w) {
    worker = null;
    workerBroken = true;
  }
  if (active?.worker === w) {
    const entry = active;
    resolveEntry(entry, workerLoadFailure());
  }
  // Queued requests never touched the failed worker: they either continue on
  // the replacement worker or, when workerBroken is set, fall back to main.
  dispatchNext();
}

function dispatchNext(): void {
  while (!active && queue.length > 0) {
    const entry = queue.shift()!;
    if (entry.req.signal?.aborted) {
      resolveEntry(entry, abortedResult());
      continue;
    }
    const w = getOrCreateWorker();
    if (!w) {
      // worker unavailable/broken -> caller's main-thread fallback.
      resolveEntry(entry, workerLoadFailure());
      continue;
    }
    entry.worker = w;
    active = entry;
    const { wasmUrl, workerTimeoutMs } = getWasmLoader().getOptions();
    // bytes 不转移（结构化克隆复制）：几 MB 文件的复制成本毫秒级，换取
    // worker 失败后主线程兜底仍持有原始字节——转移会把 buffer 中性化，
    // 兜底即失效。
    try {
      w.postMessage({
        id: entry.id,
        bytes: entry.req.bytes,
        password: entry.req.password,
        // 必须 String()：URL 对象不可结构化克隆，postMessage 会同步抛
        // DataCloneError（Chrome/Edge 实测），而默认路径正是 new URL(...)。
        wasmUrl: String(wasmUrl ?? defaultWasmUrl()),
      });
    } catch (err) {
      entry.worker = null;
      resolveEntry(entry, {
        ok: false,
        code: "UNKNOWN",
        message:
          "parse worker postMessage failed: " +
          (err instanceof Error ? err.message : String(err)),
      });
      continue;
    }

    // 计时从 postMessage 成功后开始：仅覆盖 worker 实际处理时间，不含
    // 排队等待与结构化克隆成本。超时只收尾当前请求；排队请求落到新
    // worker，不再被连坐拒绝。
    const timeoutMs = workerTimeoutMs ?? 120_000;
    entry.timer = setTimeout(() => {
      if (active !== entry) return;
      w.terminate();
      if (worker === w) worker = null;
      resolveEntry(entry, {
        ok: false,
        code: "UNKNOWN",
        message: `parse worker timed out after ${timeoutMs}ms (worker terminated); retry the parse`,
        noFallback: true,
      });
      dispatchNext();
    }, timeoutMs);
    return;
  }
}

/** 队列入口。abort 只影响本请求：排队中摘除，活动中终止 worker 并把队列
 * 重新派发到新实例。 */
function enqueue(entry: QueueEntry): void {
  if (entry.req.signal?.aborted) {
    entry.resolve(abortedResult());
    return;
  }
  if (entry.req.signal) {
    entry.onAbort = () => {
      if (entry === active) {
        const w = entry.worker;
        if (w) {
          w.terminate();
          if (worker === w) worker = null;
        }
        resolveEntry(entry, abortedResult());
        dispatchNext();
        return;
      }
      const idx = queue.indexOf(entry);
      if (idx >= 0) queue.splice(idx, 1);
      finishEntry(entry);
      entry.resolve(abortedResult());
    };
    entry.req.signal.addEventListener("abort", entry.onAbort, { once: true });
  }
  queue.push(entry);
  dispatchNext();
}

/** 拒绝全部在途/排队请求（worker 报废时整批收尾）。 */
function rejectAllPending(err: Omit<ParseErr, "ok">): void {
  const entries = active ? [active, ...queue] : [...queue];
  active = null;
  queue.length = 0;
  for (const entry of entries) resolveEntry(entry, { ok: false, ...err });
}

/**
 * 终止共享解析 worker 并拒绝全部在途/排队请求（公开逃生口，主入口导出）。
 * 在途请求标 noFallback（与超时/导出包 terminateWorker 同口径）：主动终止
 * 通常发生在换 parseWorkerUrl/wasmUrl 或页面收尾时，此后把在途的大文件解析
 * 重跑到主线程既违背终止意图、又是 worker 路由要规避的冻结源——直接失败，
 * 调用方重试即可。
 *
 * 同时清除 workerBroken：换 URL 后调用本函数，下一次解析会按 loader 当前
 * 配置重建新 worker（此前 workerBroken 会让 getOrCreateWorker 永久返回 null，
 * 文档承诺的逃生口实际失效）。不置 workerBroken 的超时路径同样在下次解析
 * 重建新实例。
 */
export function terminateParseWorker(): void {
  worker?.terminate();
  worker = null;
  workerBroken = false;
  rejectAllPending({
    code: "UNKNOWN",
    message:
      "parse worker terminated via terminateParseWorker(); retry the parse",
    noFallback: true,
  });
}

/** 主线程解析（Node/SSR/测试/Worker 失败回退共用）。 */
export async function parseOnMainThread(
  req: ParseRequest,
): Promise<ParseResult> {
  const t0 = performance.now();
  try {
    if (req.signal?.aborted) return abortedResult();
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
    // WASM 加载可能耗时数秒，等待期间中止即放弃后续解析
    if (req.signal?.aborted) return abortedResult();
    const wb = await readBuffer(
      req.bytes,
      req.password ? { password: req.password } : undefined,
    );
    // readBuffer 返回后的微任务边界：模型构建是下一段大计算（100k 行
    // ~2.6s 实测），中止即放弃（同步 WASM 段内无法打断，只能边界放弃）
    if (req.signal?.aborted) return abortedResult();
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
  // 已中止的 signal 直接短路：不构造 worker、不发消息、不跑主线程解析
  if (req.signal?.aborted) return abortedResult();
  const w = getOrCreateWorker();
  if (!w) return parseOnMainThread(req);

  const t0 = performance.now();
  const id = ++requestIdSeq;
  const result = await new Promise<ParseResult>((resolve) => {
    enqueue({ id, req, resolve, worker: null });
  });
  if (result.ok) {
    return { ...result, parseMs: performance.now() - t0 };
  }
  // 调用方取消：直传，不回退主线程（对取消的输入复跑违背取消意图）
  if (result.code === "ABORTED") return result;
  // 域错误（密码/格式判别等确定性结论）直传：主线程复跑必然同错，
  // 回退只会把加密/损坏文件的解析成本翻倍
  if (result.domain) return result;
  // worker 超时：不回退主线程（见 ParseErr.noFallback）
  if (result.noFallback) return result;
  // Worker 失败（脚本加载失败/内部错误）→ 主线程兜底（bytes 未转移，可用）
  return parseOnMainThread(req);
}
