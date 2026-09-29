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

/** 嗅探分隔符：首段文本中引号外出现频次最高者。 */
function sniffDelimiter(text: string): string {
  const sample = text.slice(0, 8192);
  const counts = new Map<string, number>();
  let inQuote = false;
  const candidates = [",", ";", "\t", "|"];
  for (const ch of sample) {
    if (ch === '"') {
      inQuote = !inQuote;
      continue;
    }
    if (inQuote) continue;
    if (candidates.includes(ch)) {
      counts.set(ch, (counts.get(ch) ?? 0) + 1);
    }
  }
  let best = ",";
  let bestCount = counts.get(",") ?? 0;
  for (const [d, n] of counts) {
    if (n > bestCount) {
      best = d;
      bestCount = n;
    }
  }
  return bestCount === 0 ? "," : best;
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
  if (Number.isFinite(n) && /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(v.trim())) {
    return { type: "number", value: v.trim() };
  }
  return { type: "string", value: v };
}

/**
 * CSV 字节 → PreviewWorkbook（单 sheet、无样式、无冻结/合并）。
 * 超 16384 列/超宽行按原样保留（渲染层有列上限保护）。
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
