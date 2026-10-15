import type { ColumnConfig, ExportOptions, ExportResult } from "./types";
import { columnLabel } from "./column-tree";
import { defaultWasmUrl, getWasmLoader } from "@marcusok/xlsx-core";
import { toBlobPart } from "./download";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

interface PendingEntry {
  resolve: (b: Uint8Array, rowCount: number) => void;
  reject: (e: Error) => void;
  /** The worker instance this request was dispatched to (see onerror). */
  worker: Worker;
  onProgress?: (progress: number) => void;
  onPhase?: (phase: "init" | "build", durationMs: number) => void;
}

let worker: Worker | null = null;
let requestIdSeq = 0;
// requestId -> pending; replaces the single-onmessage pattern that broke concurrency.
const pending = new Map<number, PendingEntry>();

interface WorkerOkResponse {
  id: number;
  ok: true;
  bytes: Uint8Array;
  rowCount: number;
  engine: "modern-xlsx" | "fast-xlsx";
}

/**
 * 超时类失败标记（exportInWorker 创建、路由层 index.ts 消费）：worker 超时
 * 意味着它已挂死（wasm 初始化悬置/引擎死循环）或负载 legitimately 超出
 * workerTimeoutMs——两条主线程重跑都避不开（同样的大负载同步阻塞主线程，
 * 正是 worker 路由要规避的冻结源，且挂死场景大概率同样挂死）。与
 * excel-preview 解析超时的 noFallback 同口径：直接失败，调用方重试即可
 * （重试会重建新 worker）。其余失败（脚本加载/克隆错误/worker 内 WASM
 * 失败）不预示主线程也会失败，仍走主线程重试，不带本标记。
 */
export type NoFallbackError = Error & { noFallback: true };

function noFallbackError(message: string): NoFallbackError {
  return Object.assign(new Error(message), { noFallback: true as const });
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

function getOrCreateWorker(): Worker {
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
  const w = new Worker(url, { type: "module" });
  worker = w;
  // Single onmessage handler registered once, dispatches by id.
  worker.onmessage = (e: MessageEvent<WorkerResponse>) => {
    const data = e.data;
    const p = pending.get(data.id);
    if (!p) return;
    // Progress messages do not complete the export; forward and keep pending.
    if ("progress" in data) {
      p.onProgress?.(data.progress);
      return;
    }
    if ("phase" in data) {
      p.onPhase?.(data.phase, data.duration);
      return;
    }
    pending.delete(data.id);
    if (data.ok && data.bytes) p.resolve(data.bytes, data.rowCount);
    else
      p.reject(
        new Error((data as WorkerErrResponse).error ?? "worker unknown error"),
      );
  };
  // A worker that errored (e.g. script failed to load) must not be reused:
  // terminate it and drop the cached reference so the next export creates a
  // fresh one, instead of failing forever into the stream fallback. Only the
  // requests dispatched to THIS worker are rejected -- a replacement worker
  // may already be serving newer request ids.
  w.onerror = (err) => {
    if (worker === w) {
      w.terminate();
      worker = null;
    }
    for (const [id, p] of pending) {
      if (p.worker !== w) continue;
      pending.delete(id);
      p.reject(new Error(err.message || "worker error"));
    }
  };
  return w;
}

/**
 * Strip non-cloneable / worker-irrelevant fields before structured clone.
 * Functions (callbacks, function-form formats) cannot be cloned; `overlay` is
 * consumed entirely by the exportExcel wrapper on the main thread (the mask is
 * already wired there) and may legally carry DOM values (e.g. `container:
 * HTMLElement`) — leaving it in the payload makes every worker-route postMessage
 * throw DataCloneError, silently degrading 20k+ row exports to a main-thread
 * rebuild the worker threshold exists to avoid.
 */
function stripFunctionFormats(options: ExportOptions): ExportOptions {
  const {
    onProgress: _onProgress,
    onPhase: _onPhase,
    overlay: _overlay,
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
  // Always forward a resolved URL: with the default (no configureWasm call)
  // the worker cannot locate the wasm next to its own bundled location —
  // the main thread's default points at the asset this bundle actually ships.
  // String() is mandatory: a URL object is not structured-cloneable, so
  // postMessage would throw DataCloneError and silently degrade every
  // default-configured browser export to the main thread.
  const wasmUrl = String(
    getWasmLoader().getOptions().wasmUrl ?? defaultWasmUrl(),
  );
  const id = ++requestIdSeq;

  try {
    const w = getOrCreateWorker();
    const timeoutMs = getWasmLoader().getOptions().workerTimeoutMs ?? 120_000; // 2-minute default
    const [bytes, workerRowCount] = await new Promise<[Uint8Array, number]>(
      (resolve, reject) => {
        const timer = setTimeout(() => {
          pending.delete(id);
          // The timed-out request can no longer receive its result. Terminate
          // the worker: either it is wedged (must not be reused by later
          // exports) or legitimately still working on a huge export whose
          // result is now unwanted — either way, keeping it burns CPU and
          // delays every request queued behind it. Sibling requests dispatched
          // to the same worker are rejected so their callers degrade through
          // the main-thread retry immediately, exactly like an onerror.
          if (worker === w) {
            w.terminate();
            worker = null;
          }
          for (const [pid, p] of pending) {
            if (p.worker !== w) continue;
            pending.delete(pid);
            // 连坐请求同样标 noFallback：共享 worker 已被 >timeoutMs 的大负载
            // 占死，排队请求的主线程重试只会叠加阻塞（见 NoFallbackError）。
            p.reject(
              noFallbackError(
                "worker terminated after a concurrent export timed out",
              ),
            );
          }
          reject(
            noFallbackError(
              `export worker timed out after ${timeoutMs}ms (worker terminated). Retry the export, raise workerTimeoutMs via configureWasm(), or split the input`,
            ),
          );
        }, timeoutMs);
        pending.set(id, {
          worker: w,
          resolve: (b: Uint8Array, rc: number) => {
            clearTimeout(timer);
            resolve([b, rc]);
          },
          reject: (e: Error) => {
            clearTimeout(timer);
            reject(e);
          },
          onProgress: options.onProgress,
          onPhase: (phase, duration) => options.onPhase?.(phase, duration),
        });
        try {
          w.postMessage({
            id,
            options: stripFunctionFormats(options),
            wasmUrl,
            mode,
          });
        } catch (e) {
          // postMessage throws synchronously on un-cloneable payloads (e.g.
          // Symbol/function values in row data -> DataCloneError). Clean up
          // immediately: otherwise the timeout timer above would fire later
          // for this id and terminate the healthy shared worker, also
          // rejecting every sibling request dispatched to it.
          pending.delete(id);
          clearTimeout(timer);
          reject(e instanceof Error ? e : new Error(String(e)));
        }
      },
    );
    const blob = new Blob([toBlobPart(bytes)], { type: XLSX_MIME });
    return {
      success: true,
      blob,
      // stream 模式在 worker 内跑的也是本包纯 JS 写出器（不经 WASM 引擎），
      // 与主线程 stream 路径同口径如实上报
      engine: mode === "stream" ? "fast-xlsx" : "modern-xlsx",
      mode: mode === "stream" ? "stream" : "worker",
      duration: performance.now() - start,
      rowCount: workerRowCount,
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
 * 终止共享导出 worker 并拒绝全部在途请求。在途请求标 noFallback（与超时
 * 连坐同口径，见 NoFallbackError）：主动终止通常发生在换 wasmUrl/workerUrl
 * 或页面收尾时，此后路由层把同一大负载重跑到主线程既违背终止意图、又会
 * 把 worker 路由要规避的主线程冻结原样补回来——直接失败，调用方重试即可
 * （重试会以新配置重建 worker）。
 */
export function terminateWorker(): void {
  worker?.terminate();
  worker = null;
  for (const [, p] of pending)
    p.reject(
      noFallbackError(
        "export worker terminated via terminateWorker(); retry the export",
      ),
    );
  pending.clear();
}
