/**
 * 引擎错误 → 预览错误码的统一归一（worker 与主线程回退两条路径共用）。
 *
 * 此前 worker 内与 worker-client 各写一份近似逻辑，engine 的
 * UNRECOGNIZED_FORMAT 在 worker 路径被原样透传——它不在公共类型
 * PreviewErrorCode 联合内，且与主线程路径（归一为 CORRUPT/LEGACY_FORMAT）
 * 对同一份字节给出不同 code。收敛到本模块后两条路径结论恒一致。
 *
 * 2026-09 补齐结构类与 WASM 类码的映射：此前只有 4 个码被识别，其余全部
 * 落到 UNKNOWN——实测 .ods/.docx 同形 zip（MISSING_PART）、截断的 xlsx
 * （ZIP_READ）拿到的是 UNKNOWN 加引擎原文，而文档承诺的是 CORRUPT；WASM
 * 加载失败同样是 UNKNOWN，调用方无法按码分流。
 */
import {
  LEGACY_FORMAT,
  PASSWORD_PROTECTED,
  UNRECOGNIZED_FORMAT,
  WASM_INIT_FAILED,
  type ModernXlsxError,
} from "@marcusok/xlsx-core";
import type { PreviewErrorCode } from "../types";

export interface EngineErrorInfo {
  code: PreviewErrorCode;
  message: string;
  /**
   * true = 对相同字节必然复现的域错误（格式/密码判别）：worker 收到后直传
   * 调用方，不回退主线程复跑（复跑必然同错，大文件在主线程整段重跑正是
   * 回退机制要规避的冻结源）。
   */
  domain: boolean;
}

const LEGACY_MSG =
  "Legacy .xls (BIFF8) workbooks are not supported. Re-save the file as .xlsx and try again.";
const CORRUPT_MSG =
  "Not a valid .xlsx/.xlsm file (ZIP structure expected). The file may be corrupt or in an unsupported format.";
const PASSWORD_MSG =
  "The workbook is password-protected (Agile AES-256). Pass the `password` option to open it.";
const WASM_MSG =
  "The WebAssembly engine is unavailable or failed to load (a blocked/404 asset URL, a CSP that forbids WebAssembly, or an environment without it). Check the configured `wasmUrl` and the console.";

/**
 * 结构类错误码（引擎读路径实测抛出，2026-09）：ZIP 打不开（ZIP_READ）、
 * 局部目录损坏（ZIP_ENTRY）、缺必需部件（MISSING_PART：非 xlsx 的 zip
 * ——.ods/.docx 同形——与缺 sheet 的残包）、部件 XML 解析失败（XML_PARSE）。
 * 这些都表示"字节不是有效 xlsx"，归一为 CORRUPT 的友好文案。
 *
 * modern-xlsx 未在 xlsx-core 的再导出面内提供这些常量（其 re-export 只有
 * LEGACY_FORMAT / UNRECOGNIZED_FORMAT / PASSWORD_PROTECTED / WASM_INIT_FAILED
 * 四个），按字面值匹配；码值是引擎类型层 ErrorCode 联合的一部分，属稳定契约。
 */
const STRUCTURAL_CODES = new Set([
  "ZIP_READ",
  "ZIP_ENTRY",
  "MISSING_PART",
  "XML_PARSE",
]);

/** WASM 环境类错误码（WASM_INIT_FAILED 有导出的常量，其余按字面值）。 */
const WASM_CODES = new Set(["WASM", "WASM_ERROR", WASM_INIT_FAILED]);

/**
 * 预览自产错误码（模型/解析层抛出的、已经归一过的码，如零 sheet 的
 * CORRUPT）：透传，避免被下面的映射再折成 UNKNOWN。
 */
const PASSTHROUGH_CODES = new Set<PreviewErrorCode>(["CORRUPT", "UNSUPPORTED"]);

export function normalizeEngineError(err: unknown): EngineErrorInfo {
  const rawCode = (err as ModernXlsxError)?.code;
  if (typeof rawCode !== "string") {
    return {
      code: "UNKNOWN",
      message: err instanceof Error ? err.message : String(err),
      domain: false,
    };
  }
  // WASM 失败属环境性而非文件性：不标 domain，允许回退主线程复跑
  // （worker 的 CSP 与主线程可能不同，worker 内失败不代表主线程也失败）。
  if (WASM_CODES.has(rawCode)) {
    return { code: "WASM", message: WASM_MSG, domain: false };
  }
  if (rawCode === PASSWORD_PROTECTED) {
    return { code: "PASSWORD_PROTECTED", message: PASSWORD_MSG, domain: true };
  }
  if (rawCode === LEGACY_FORMAT) {
    return { code: "LEGACY_FORMAT", message: LEGACY_MSG, domain: true };
  }
  if (rawCode === UNRECOGNIZED_FORMAT) {
    // .xls（BIFF8）判别双码兜底：真实/伪造 OLE2 复合文档实测抛
    // UNRECOGNIZED_FORMAT，真实复合文档预期抛 LEGACY_FORMAT——message 文本
    // 随输入形态变化（实测），判别依据只有 code 与是否含 OLE2 字样。
    const message = err instanceof Error ? err.message : String(err);
    return /ole2/i.test(message)
      ? { code: "LEGACY_FORMAT", message: LEGACY_MSG, domain: true }
      : { code: "CORRUPT", message: CORRUPT_MSG, domain: true };
  }
  // 结构类错误：对相同字节确定性成立，标 domain 免去主线程复跑
  if (STRUCTURAL_CODES.has(rawCode)) {
    return { code: "CORRUPT", message: CORRUPT_MSG, domain: true };
  }
  if (PASSTHROUGH_CODES.has(rawCode as PreviewErrorCode)) {
    return {
      code: rawCode as PreviewErrorCode,
      message: err instanceof Error ? err.message : String(err),
      domain: true,
    };
  }
  return {
    code: "UNKNOWN",
    message: err instanceof Error ? err.message : String(err),
    domain: false,
  };
}
