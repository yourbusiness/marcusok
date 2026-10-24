import type { ColumnConfig, ExportOptions, ExportResult } from "./types";
import { columnLabel } from "./column-tree";
import { defaultWasmUrl, getWasmLoader } from "@marcusok/xlsx-core";
import { toBlobPart } from "./download";
import { streamDroppedFeatures } from "./stream-features";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/**
 * 单 worker + 显式 FIFO 队列。
 *
 * 原实现把每个请求直接 postMessage 给共享 worker，并用"派发即开始"的定时器
 * 计时：队列里没开始处理的请求也被计入超时，一个长任务就能让后面本可成功的
 * 小请求超时，并连坐拒绝所有在途请求。现在同一时刻只派发一个请求，计时从
 * postMessage 成功后开始（只覆盖 worker 实际处理时间）；超时/中止只收尾当前
 * 请求，排队的请求继续按序落到新 worker。并发导出仍然串行（共享 worker 的
 * 固有限制），但不会再互相误伤。
 */
interface QueueEntry {
  id: number;
  options: ExportOptions;
  mode: "workbook" | "stream";
  timeoutMs: number;
  resolve: (b: Uint8Array, rowCount: number) => void;
  reject: (e: Error) => void;
  onProgress?: (progress: number) => void;
  onPhase?: (phase: "init" | "build", durationMs: number) => void;
  timer?: ReturnType<typeof setTimeout>;
  worker: Worker | null;
  signal?: AbortSignal;
  onAbort?: () => void;
}

let worker: Worker | null = null;
// 同步构造失败（如跨域 workerUrl）置位后不再重复构造；换 URL 后须先调用
// terminateWorker() 清除该标志，才能按新配置重建。
let workerBroken = false;
let requestIdSeq = 0;
let active: QueueEntry | null = null;
const queue: QueueEntry[] = [];

interface WorkerOkResponse {
  id: number;
  ok: true;
  bytes: Uint8Array;
  rowCount: number;
  engine: "modern-xlsx" | "fast-xlsx";
}

/**
 * 超时类失败标记（exportInWorker 创建、路由层 index.ts 消费）：worker 超时
 * 意味着它已挂死（wasm 初始化悬置/引擎死循环）或单次负载 legitimately 超出
 * workerTimeoutMs——主线程重跑都避不开（同样的大负载同步阻塞主线程，正是
 * worker 路由要规避的冻结源，且挂死场景大概率同样挂死）。与 excel-preview
 * 解析超时的 noFallback 同口径：直接失败，调用方重试即可（重试会重建新
 * worker）。其余失败（脚本加载/克隆错误/worker 内 WASM 失败）不预示主线程
 * 也会失败，仍走主线程重试，不带本标记。
 */
export type NoFallbackError = Error & { noFallback: true };

function noFallbackError(message: string): NoFallbackError {
  return Object.assign(new Error(message), { noFallback: true as const });
}

/** 调用方通过 AbortSignal 主动取消；路由层据此跳过降级链。 */
function abortError(): Error {
  const err = new Error("[excel-exporter] export aborted by the caller");
  err.name = "AbortError";
  return err;
}

export function isAbortError(e: unknown): e is Error {
  return e instanceof Error && e.name === "AbortError";
}

interface WorkerErrResponse {
  id: number;
  ok: false;
  error: string;
}
interface WorkerProgressResponse {
  id: number;
  progress: number;
}
interface WorkerPhaseResponse {
  id: number;
  phase: "init" | "build";
  duration: number;
}
type WorkerResponse =
  | WorkerOkResponse
  | WorkerErrResponse
  | WorkerProgressResponse
  | WorkerPhaseResponse;

/** 收尾当前请求的计时器/abort 监听（不负责 resolve/reject）。 */
function finishEntry(entry: QueueEntry): void {
  if (entry.timer !== undefined) {
    clearTimeout(entry.timer);
    entry.timer = undefined;
  }
  if (entry.signal && entry.onAbort) {
    entry.signal.removeEventListener("abort", entry.onAbort);
    entry.onAbort = undefined;
  }
  if (active === entry) active = null;
}

function rejectEntry(entry: QueueEntry, err: Error): void {
  finishEntry(entry);
  entry.reject(err);
}

function handleMessage(w: Worker, data: WorkerResponse): void {
  // Responses from a replaced/terminated worker must not settle the request
  // now running on its replacement.
  if (!active || active.worker !== w || active.id !== data.id) return;
  const entry = active;
  // Progress/phase messages do not complete the export.
  if ("progress" in data) {
    entry.onProgress?.(data.progress);
    return;
  }
  if ("phase" in data) {
    entry.onPhase?.(data.phase, data.duration);
    return;
  }
  finishEntry(entry);
  if (data.ok && data.bytes) entry.resolve(data.bytes, data.rowCount);
  else
    entry.reject(
      new Error((data as WorkerErrResponse).error ?? "worker unknown error"),
    );
  dispatchNext();
}

function handleError(w: Worker, err: Error): void {
  if (worker === w) {
    w.terminate();
    worker = null;
  }
  // Only the request actually dispatched to this worker is settled. Queued
  // requests were never sent to it; dispatchNext gives them a fresh worker.
  if (active?.worker === w) {
    rejectEntry(active, new Error(err.message || "worker error"));
  }
  dispatchNext();
}

function getOrCreateWorker(): Worker {
  if (workerBroken) {
    throw new Error(
      "export worker construction previously failed; call terminateWorker() before retrying with a new workerUrl",
    );
  }
  if (worker) return worker;
  const { workerUrl } = getWasmLoader().getOptions();
  // Hoisted deliberately. The inline `new Worker(new URL(<literal>,
  // import.meta.url), ...)` form makes bundlers RE-BUNDLE the worker as a
  // build entry — Vite 5 (VitePress) even fails outright on it (its default
  // iife worker format rejects the code-splitting our engine's dead Node
  // dynamic imports produce). Hoisted, only the asset-import-meta-url pattern
  // matches, so bundlers emit the shipped file verbatim as a hashed asset —
  // and dist/export.worker.js is a self-contained ESM with zero imports,
  // built precisely so the verbatim copy loads as-is.
  const url = workerUrl ?? new URL("./export.worker.js", import.meta.url);
  let w: Worker;
  try {
    w = new Worker(url, { type: "module" });
  } catch (err) {
    // 跨域脚本会在构造时同步抛 SecurityError。标记后让后续导出直接走主线程
    // 降级，避免固定错误配置在每次导出时重复构造；原始异常继续交给路由层。
    workerBroken = true;
    console.warn(
      "[excel-exporter] export worker failed to construct: " +
        (err instanceof Error ? err.message : String(err)) +
        ". Falling back to main-thread export. Worker scripts must be same-origin — serve workerUrl from your own origin; only the WASM binary may be cross-origin.",
    );
    throw err;
  }
  worker = w;
  w.onmessage = (e: MessageEvent<WorkerResponse>) => handleMessage(w, e.data);
  w.onerror = (err) => handleError(w, new Error(err.message || "worker error"));
  return w;
}

function dispatchNext(): void {
  while (!active && queue.length > 0) {
    const entry = queue.shift()!;
    if (entry.signal?.aborted) {
      rejectEntry(entry, abortError());
      continue;
    }

    let w: Worker;
    try {
      w = getOrCreateWorker();
    } catch (e) {
      rejectEntry(entry, e instanceof Error ? e : new Error(String(e)));
      continue;
    }

    entry.worker = w;
    active = entry;
    try {
      w.postMessage({
        id: entry.id,
        options: stripFunctionFormats(entry.options),
        wasmUrl: String(
          getWasmLoader().getOptions().wasmUrl ?? defaultWasmUrl(),
        ),
        mode: entry.mode,
      });
    } catch (e) {
      // postMessage throws synchronously on un-cloneable payloads. Keep the
      // healthy worker for the next request; this entry fails immediately
      // without a dangling timer.
      entry.worker = null;
      rejectEntry(entry, e instanceof Error ? e : new Error(String(e)));
      continue;
    }

    // Start the clock after postMessage: structured-clone cost is main-thread
    // work, not worker processing time. Queue wait is excluded by design.
    entry.timer = setTimeout(() => {
      if (active !== entry) return;
      const timedOutWorker = entry.worker;
      if (timedOutWorker && worker === timedOutWorker) {
        timedOutWorker.terminate();
        worker = null;
      }
      const timeoutMs = entry.timeoutMs;
      rejectEntry(
        entry,
        noFallbackError(
          `export worker timed out after ${timeoutMs}ms (worker terminated). Retry the export, raise workerTimeoutMs via configureWasm(), or split the input`,
        ),
      );
      dispatchNext();
    }, entry.timeoutMs);
    return;
  }
}

function enqueue(entry: QueueEntry): void {
  if (entry.signal?.aborted) {
    entry.reject(abortError());
    return;
  }
  if (entry.signal) {
    entry.onAbort = () => {
      if (entry === active) {
        // Terminate the active worker: the request can no longer receive its
        // result, and the queued requests are re-dispatched to a fresh one.
        const w = entry.worker;
        if (w && worker === w) {
          w.terminate();
          worker = null;
        }
        rejectEntry(entry, abortError());
        dispatchNext();
        return;
      }
      const idx = queue.indexOf(entry);
      if (idx >= 0) queue.splice(idx, 1);
      entry.reject(abortError());
    };
    entry.signal.addEventListener("abort", entry.onAbort, { once: true });
  }
  queue.push(entry);
  dispatchNext();
}

/**
 * Strip non-cloneable / worker-irrelevant fields before structured clone.
 * Functions (callbacks, function-form formats) cannot be cloned; `overlay` is
 * consumed entirely by the exportExcel wrapper on the main thread (the mask is
 * already wired there) and may legally carry DOM values (e.g. `container:
 * HTMLElement`); `signal` is an AbortSignal (not structured-cloneable) and is
 * consumed on the main thread. Leaving them in the payload makes every
 * worker-route postMessage throw DataCloneError, silently degrading 20k+ row
 * exports to a main-thread rebuild the worker threshold exists to avoid.
 */
function stripFunctionFormats(options: ExportOptions): ExportOptions {
  const {
    onProgress: _onProgress,
    onPhase: _onPhase,
    overlay: _overlay,
    signal: _signal,
    ...rest
  } = options;
  return {
    ...rest,
    sheets: options.sheets.map((s) => ({
      ...s,
      columns: s.columns.map(stripColumn),
    })),
  };
}

/** Recursively strip function-form formats; group columns carry `children`. */
function stripColumn(c: ColumnConfig): ColumnConfig {
  const children = c.children?.map(stripColumn);
  if (c.format && typeof c.format === "function") {
    console.warn(
      `[excel-exporter] column "${columnLabel(c)}" uses a function format, stripped for worker mode. Use FormatSpec for worker compatibility.`,
    );
    const { format: _format, ...rest } = c;
    return children ? { ...rest, children } : rest;
  }
  return children ? { ...c, children } : c;
}

export async function exportInWorker(
  options: ExportOptions,
  mode: "workbook" | "stream",
): Promise<ExportResult> {
  const start = performance.now();
  const id = ++requestIdSeq;
  const timeoutMs = getWasmLoader().getOptions().workerTimeoutMs ?? 120_000;

  try {
    const [bytes, workerRowCount] = await new Promise<[Uint8Array, number]>(
      (resolve, reject) => {
        enqueue({
          id,
          options,
          mode,
          timeoutMs,
          resolve: (bytes, rowCount) => resolve([bytes, rowCount]),
          reject,
          onProgress: options.onProgress,
          onPhase: (phase, duration) => options.onPhase?.(phase, duration),
          worker: null,
          signal: options.signal,
        });
      },
    );
    const blob = new Blob([toBlobPart(bytes)], { type: XLSX_MIME });
    const droppedFeatures =
      mode === "stream" ? streamDroppedFeatures(options.sheets) : [];
    return {
      success: true,
      blob,
      // stream 模式在 worker 内跑的也是本包纯 JS 写出器（不经 WASM 引擎），
      // 与主线程 stream 路径同口径如实上报
      engine: mode === "stream" ? "fast-xlsx" : "modern-xlsx",
      mode: mode === "stream" ? "stream" : "worker",
      duration: performance.now() - start,
      rowCount: workerRowCount,
      ...(droppedFeatures.length > 0 && { degradedFeatures: droppedFeatures }),
    };
  } catch (e) {
    return {
      success: false,
      error: e as Error,
      duration: performance.now() - start,
    };
  }
}

/**
 * 终止共享导出 worker 并拒绝全部在途/排队请求。在途请求标 noFallback（与
 * 超时同口径，见 NoFallbackError）：主动终止通常发生在换 wasmUrl/workerUrl
 * 或页面收尾时，此后路由层把同一大负载重跑到主线程既违背终止意图、又会
 * 把 worker 路由要规避的主线程冻结原样补回来——直接失败，调用方重试即可
 * （重试会以新配置重建 worker）。排队请求同样失败：主动终止是全局意图，
 * 不应留下继续执行的请求。
 */
export function terminateWorker(): void {
  worker?.terminate();
  worker = null;
  workerBroken = false;
  const entries = active ? [active, ...queue] : [...queue];
  active = null;
  queue.length = 0;
  for (const entry of entries) {
    finishEntry(entry);
    entry.reject(
      noFallbackError(
        "export worker terminated via terminateWorker(); retry the export",
      ),
    );
  }
}
