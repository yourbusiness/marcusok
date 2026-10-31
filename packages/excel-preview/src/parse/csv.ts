/**
 * CSV 轻量解析：无样式、无合并，产出单 sheet 的 PreviewWorkbook（复用同一
 * 渲染管线）。编码 UTF-8（BOM 剥离）→ GB18030 回退；分隔符嗅探（, ; \t |）；
 * RFC 4180 引号转义合容。
 */
import type { PreviewWorkbook, PreviewSheet } from "../types";

const DEFAULT_STYLES = {
  fonts: [],
  fills: [],
  borders: [],
  xfs: [],
};

/**
 * 嗅探分隔符：按"每行字段数的一致性"评分，而不是字符总频次。
 *
 * 纯频次会被字段内容投票：TSV 的文本列里逗号常多于制表符
 * （`Widget, large, red, blue\t10`），整表会被按逗号拆错列且无告警。
 * 正解是候选分隔符各自试解析样本，取"众数字段数 × 众数行占比"最高者：
 * 真正的分隔符让每一行（含单列行）落在同一列数上，正文里的偶发符号
 * 只会制造少量离群行。引用内的分隔符由 parseCsvText 正确跳过。
 */
function sniffDelimiter(text: string): string {
  const sample = text.slice(0, 8192);
  const candidates = [",", ";", "\t", "|"];
  let best = ",";
  let bestScore = 0;
  for (const delimiter of candidates) {
    const rows = parseCsvText(sample, delimiter).filter(
      (r) => !(r.length === 1 && r[0] === ""),
    );
    if (rows.length === 0) continue;
    const counts = new Map<number, number>();
    for (const row of rows) {
      counts.set(row.length, (counts.get(row.length) ?? 0) + 1);
    }
    let modalCols = 0;
    let modalRows = 0;
    for (const [cols, n] of counts) {
      // 同票取列数更多者：一致的两列结构比一致的单列更可能是分隔符。
      if (n > modalRows || (n === modalRows && cols > modalCols)) {
        modalCols = cols;
        modalRows = n;
      }
    }
    const score = (modalCols * modalRows) / rows.length;
    // 严格大于：平分时保留候选顺序（逗号优先），与旧默认一致。
    if (score > bestScore) {
      best = delimiter;
      bestScore = score;
    }
  }
  return best;
}

/** RFC 4180 解析（支持 CRLF/LF、双引号转义、跨行引号字段）。 */
export function parseCsvText(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuote = false;
  let i = 0;
  const n = text.length;
  while (i < n) {
    const ch = text[i];
    if (inQuote) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuote = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === "") {
      inQuote = true;
      i++;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = "";
      i++;
      continue;
    }
    if (ch === "\n" || ch === "\r") {
      // CRLF 只收一行
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      field = "";
      rows.push(row);
      row = [];
      i++;
      continue;
    }
    field += ch;
    i++;
  }
  if (field !== "" || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  // 尾部空行剔除（文件以换行结尾的常态）
  while (rows.length > 0) {
    const last = rows[rows.length - 1];
    if (last.length === 1 && last[0] === "") rows.pop();
    else break;
  }
  return rows;
}

/** 字节 → 文本：UTF-8 优先（fatal 探测），失败回退 GB18030（中文 CSV 常见）。 */
function decodeBytes(bytes: Uint8Array): string {
  // UTF-8 BOM 剥离
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder("utf-8").decode(bytes.subarray(3));
  }
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    try {
      return new TextDecoder("gb18030").decode(bytes);
    } catch {
      return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
    }
  }
}

/** 类型推断单元格（Excel 的 General 语义：数字右对齐，文本左对齐）。 */
function csvCellType(v: string): { type: "number" | "string"; value: string } {
  if (v === "") return { type: "string", value: "" };
  const n = Number(v);
  // \.\d* 而非 \.\d+：尾点数字（"123."）在 Excel 中即数字 123，点后
  // 允许无数字；前点（".5"）与多段数字（"1.2.3"）仍不匹配，留在文本类
  if (Number.isFinite(n) && /^-?\d+(\.\d*)?([eE][+-]?\d+)?$/.test(v.trim())) {
    return { type: "number", value: v.trim() };
  }
  return { type: "string", value: v };
}

/**
 * CSV 字节 → PreviewWorkbook（单 sheet、无样式、无冻结/合并）。
 * 注意口径：超 Excel 上限的宽行/长行会被渲染层静默截断——layout.ts 的
 * buildLayout 把 colCount/rowCount 封顶到 16384/1048576，越界格不渲染、
 * 无告警（xlsx 路径的同类输入在模型层直接拒绝为 CORRUPT，两条路径口径
 * 不一致，属已知取舍）。
 */
export function csvToWorkbook(bytes: Uint8Array): PreviewWorkbook {
  const text = decodeBytes(bytes);
  const delimiter = sniffDelimiter(text);
  const grid = parseCsvText(text, delimiter);

  let maxCol = 0;
  const rows = grid.map((cells, ri) => {
    if (cells.length > maxCol) maxCol = cells.length;
    return {
      index: ri + 1,
      height: null,
      hidden: false,
      cells: cells
        .map((v, ci) => ({ ...csvCellType(v), col: ci, styleIndex: null }))
        .filter((c) => c.value !== ""),
    };
  });

  const sheet: PreviewSheet = {
    name: "CSV",
    visible: true,
    showGridLines: true,
    rightToLeft: false,
    rowCount: rows.length,
    colCount: maxCol,
    rows,
    colSpans: [],
    merges: [],
    frozenRows: 0,
    frozenCols: 0,
    // xfs 空：渲染层对无样式模型全部走默认样式
    styles: DEFAULT_STYLES,
  };
  return {
    sheets: [sheet],
    activeSheetIndex: 0,
    dateSystem: "date1900",
  };
}
