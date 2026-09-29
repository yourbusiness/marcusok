/**
 * 解析 worker：单文件自包含（构建时 core + modern-xlsx + fflate 全部打进
 * 本文件——browser module worker 不能解析裸导入，与 excel-exporter 的
 * export.worker 同一约束与同一构建模式）。
 *
 * 协议：{id, bytes, password?, wasmUrl?} → {id, ok, workbook} |
 * {id, ok:false, code, message, domain?}。domain=true 标记确定性域错误
 * （格式/密码判别等对相同字节必然复现的结论），客户端直传不再回退主线程
 * 复跑；缺省/ false 视为环境性失败（wasm 加载、未知异常），仍可回退。
 * 模型构建（含样式覆盖层）全部在 worker 内完成，主线程拿到即渲染。
 */
import { initWasm, readBuffer } from "@marcusok/xlsx-core";
import { buildPreviewWorkbook } from "../parse/model";
import { csvToWorkbook } from "../parse/csv";
import { normalizeEngineError } from "../parse/engine-errors";
import { sniffFormat } from "../parse/sniff";

interface WorkerRequest {
  id: number;
  bytes: Uint8Array;
  password?: string;
  wasmUrl?: string | URL;
}

type WorkerResponse =
  | { id: number; ok: true; workbook: ReturnType<typeof buildPreviewWorkbook> }
  | { id: number; ok: false; code: string; message: string; domain?: boolean };

/** 见 export.worker.ts 的 loadedWasmKey 说明：initWasm 幂等，按字符串判重。 */
let wasmReady = false;
let loadedWasmKey: string | null = null;

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const { id, bytes, password, wasmUrl } = e.data;
  try {
    const form = sniffFormat(bytes);
    if (form === "text") {
      const workbook = csvToWorkbook(bytes);
      (self as unknown as Worker).postMessage({
        id,
        ok: true,
        workbook,
      } satisfies WorkerResponse);
      return;
    }
    if (form === "unknown") {
      (self as unknown as Worker).postMessage({
        id,
        ok: false,
        code: "UNSUPPORTED",
        message:
          "Unrecognized file format: expected an .xlsx/.xlsm (ZIP) or .csv (text) file.",
        // 嗅探对相同字节确定性成立，主线程复跑结论相同
        domain: true,
      } satisfies WorkerResponse);
      return;
    }

    const urlKey = wasmUrl == null ? null : String(wasmUrl);
    if (!wasmReady || loadedWasmKey !== urlKey) {
      await initWasm(wasmUrl);
      wasmReady = true;
      loadedWasmKey = urlKey;
    }

    const wb = await readBuffer(bytes, password ? { password } : undefined);
    const workbook = buildPreviewWorkbook(wb, bytes);
    (self as unknown as Worker).postMessage({
      id,
      ok: true,
      workbook,
    } satisfies WorkerResponse);
  } catch (err) {
    // 错误归一与主线程回退路径共用同一实现（parse/engine-errors）：code 直出
    // PreviewErrorCode 联合内的值（此前 UNRECOGNIZED_FORMAT 原样透传，泄漏到
    // 公共类型之外且与主线程路径结论不一致），domain 标记确定性域错误。
    const info = normalizeEngineError(err);
    (self as unknown as Worker).postMessage({
      id,
      ok: false,
      code: info.code,
      message: info.message,
      domain: info.domain,
    } satisfies WorkerResponse);
  }
};
