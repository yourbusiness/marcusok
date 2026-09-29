/**
 * 引擎错误 → 预览错误码的统一归一（worker 与主线程回退两条路径共用）。
 *
 * 此前 worker 内与 worker-client 各写一份近似逻辑，engine 的
 * UNRECOGNIZED_FORMAT 在 worker 路径被原样透传——它不在公共类型
 * PreviewErrorCode 联合内，且与主线程路径（归一为 CORRUPT/LEGACY_FORMAT）
 * 对同一份字节给出不同 code。收敛到本模块后两条路径结论恒一致。
 */
import {
  LEGACY_FORMAT,
  PASSWORD_PROTECTED,
  UNRECOGNIZED_FORMAT,
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

export function normalizeEngineError(err: unknown): EngineErrorInfo {
  const rawCode = (err as ModernXlsxError)?.code;
  if (typeof rawCode !== "string") {
    return {
      code: "UNKNOWN",
      message: err instanceof Error ? err.message : String(err),
      domain: false,
    };
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
  return {
    code: "UNKNOWN",
    message: err instanceof Error ? err.message : String(err),
    domain: false,
  };
}
