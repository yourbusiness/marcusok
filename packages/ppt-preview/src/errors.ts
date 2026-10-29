/**
 * 上游异常归一为公开错误码。
 *
 * 上游错误面（docs/ spike 记录核实）：
 * - `OoxmlError`：8 个已知 code（encrypted / invalid-password /
 *   unsupported-encryption / legacy-binary-format / not-ooxml /
 *   unsupported-image-effect / invalid-numbering / unsupported-content），
 *   只有 `code` 与可选 `details`，**没有 stage 字段**——`not-ooxml` 无法
 *   按 stage 细分损坏类型，不解析 message 强行细分（上游 message 面向
 *   人读，跨版本不稳定）；
 * - `OoxmlResourceLimitError`（code 'ooxml-resource-limit'，details.stage +
 *   violation.limit/observed）：合法文件超资源预算，不是损坏 → UNSUPPORTED；
 * - `OoxmlDecodedImageLimitError`（RangeError，code 'ooxml-decoded-image-limit'，
 *   metric/limit/observed）：解码图像超限，同上 → UNSUPPORTED；
 * - Worker/WASM/CSP/资产加载失败没有稳定 code → UNKNOWN（不误报文件损坏）。
 *
 * vi.mock 场景 instanceof 失效（pdf-preview 同款经验）：识别一律用
 * `code` 属性 + Error.name 兜底。
 */
import type { PptPreviewError } from "./types";

/** 上游资源限制错误的 code 标识（非 OoxmlError 家族，独立识别）。 */
const RESOURCE_LIMIT_CODE = "ooxml-resource-limit";
const IMAGE_LIMIT_CODE = "ooxml-decoded-image-limit";

interface ErrLike {
  code?: unknown;
  name?: string;
  message?: string;
}

function errorName(err: unknown): string {
  return err instanceof Error ? err.name : "";
}

function errorCode(err: unknown): string | undefined {
  const code = (err as ErrLike)?.code;
  return typeof code === "string" ? code : undefined;
}

/** 资源限制类：message 带 limit/observed 详情，cause 保留原始错误。 */
function limitError(err: unknown, label: string): PptPreviewError {
  const message = err instanceof Error ? err.message : String(err);
  return {
    code: "UNSUPPORTED",
    message:
      `The file exceeds a ${label} budget and cannot be previewed safely. ` +
      `Details: ${message}`,
    cause: err,
  };
}

/** 解析/渲染失败归一为公开错误码（sniff 之后的失败路径共用）。 */
export function normalizeError(err: unknown): PptPreviewError {
  const code = errorCode(err);
  const name = errorName(err);

  // 两族资源限制错误（RangeError / Error，均带固定 code）
  if (code === RESOURCE_LIMIT_CODE || name === "OoxmlResourceLimitError") {
    return limitError(err, "archive resource");
  }
  if (code === IMAGE_LIMIT_CODE || name === "OoxmlDecodedImageLimitError") {
    return limitError(err, "decoded image");
  }

  switch (code) {
    case "legacy-binary-format":
      return {
        code: "LEGACY_FORMAT",
        message:
          "Legacy binary .ppt files are not supported. Re-save the presentation as .pptx (Office Open XML).",
        cause: err,
      };
    case "encrypted":
      return {
        code: "PASSWORD_PROTECTED",
        message:
          "The file is password-protected (MS-OFFCRYPTO). Pass options.password to decrypt it.",
        cause: err,
      };
    case "invalid-password":
      return {
        code: "PASSWORD_PROTECTED",
        message:
          "The file is password-protected and the supplied password was rejected.",
        cause: err,
      };
    case "unsupported-encryption":
    case "unsupported-content":
    case "unsupported-image-effect":
      // 密码正确也解决不了的不支持项
      return {
        code: "UNSUPPORTED",
        message: err instanceof Error ? err.message : String(err),
        cause: err,
      };
    case "not-ooxml":
      // 覆盖非 PPTX、不可读容器及部分加密容器结构损坏；无 stage、message
      // 不稳定，不强行细分损坏原因（方案四节的口径）
      return {
        code: "UNSUPPORTED",
        message:
          "Not a .pptx presentation: the package could not be recognized as an OOXML PowerPoint document. " +
          `Details: ${err instanceof Error ? err.message : String(err)}`,
        cause: err,
      };
    case "invalid-numbering":
      // S2 实证可稳定区分的结构损坏
      return {
        code: "CORRUPT",
        message:
          "The .pptx file is corrupted (invalid numbering structure was found).",
        cause: err,
      };
    default:
      // 源读取、Worker、WASM、CSP、资产 URL、未知 parser trap 与未来新增
      // 的上游错误码——不误报文件损坏
      return {
        code: "UNKNOWN",
        message: `Failed to load or render the presentation: ${err instanceof Error ? err.message : String(err)}`,
        cause: err,
      };
  }
}
