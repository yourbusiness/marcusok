/**
 * 数字格式化补偿层：把 modern-xlsx formatCellRich 的输出修到接近 Excel 实际渲染。
 *
 * 引擎缺陷清单（全部 2026-09 实测复现，行为锁定为补偿依据）：
 *  1. 多段格式的负段正确选用（0.00;0.000 → "1234.500" 三位小数）但丢弃段内
 *     字面量："-"/"("（负号/会计括号）、引号串（"neg"）、货币符号（¥/$/€）、
 *     空格——数字段内的字面量全部不渲染；
 *  2. 单段格式负号正常（-1234.5 + "#,##0.00" → "-1,234.50"）；零段与文本段
 *     的引号字面量正常（"ZERO"/"零"）；日期段字面量正常（yyyy"年"m"月" ✓）；
 *  3. [h]/[m]/[s] 累计时长段全坏（"[h]:mm:ss" → ":01:00"）；无括号的
 *     mm:ss 会被当"月:秒"解析（0.25 → "12:00"，Dec 1899 的月份）；
 *  4. General 输出 16 位浮点尾迹（1/3 → "0.3333333333333333"），Excel 为
 *     15 位有效数字；
 *  5. formatCell/formatCellRich 对 null/undefined 码抛 TypeError。
 *
 * 补偿策略：
 *  - 数字段拆出"可安全回注的字面前缀/后缀"（引号串、转义符、非记号字符），
 *    中段交给引擎（以 |值| 传入，规避单段负号干扰），再拼回前后缀——负号与
 *    会计括号本身就是段首/段尾字面量，被该机制自然覆盖；
 *  - 累计时长段（含 [h]/[m]/[s] 括号记号）自实现；无括号 mm:ss 是当日
 *    时间的"分:秒"（Excel 语义；累计分钟是带括号 [m]:ss 的语义），随
 *    日期段走自实现路径（分钟邻接消歧），不进时长分支；
 *  - General 自实现 15 位有效数字；
 *  - 日期段原样交引擎（字面量无损），date1904 系统先做 +1462 天序列号平移。
 */
import { formatCellRich, serialToDate } from "@marcusok/xlsx-core";
import { colorCodeToCss } from "../color";
import type { PreviewCellType } from "../types";

export interface FormattedValue {
  text: string;
  /** numFmt 颜色段（[Red] / [Color 3]）解析出的 CSS 色。 */
  color?: string;
}

type ConditionOp = ">=" | ">" | "<=" | "<" | "=" | "<>";

interface Section {
  /** 原始段码（未去色）。 */
  raw: string;
  /** 去掉颜色/条件括号后的段码。 */
  code: string;
  color: string | null;
  condition: { op: ConditionOp; value: number } | null;
}

// 颜色括号只认已知命名色与 [Color n]——宽匹配 [a-z]+ 会把累计时长记号
// [h]/[m]/[s] 误吞成颜色（实测踩坑），时段记号必须留在码内交时长路径。
const COLOR_NAMES = "black|blue|cyan|green|magenta|red|white|yellow";
const COLOR_BRACKET = new RegExp(
  `^\\[(?:color\\s+\\d+|${COLOR_NAMES})\\]$`,
  "i",
);
const CONDITION_BRACKET = /^\[(>=|<=|<>|>|<|=)\s*(-?\d+(?:\.\d+)?)\]$/i;

/** 拆分格式码为段（引号/括号/转义内的 ";" 不拆）。 */
export function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = "";
  let inQuote = false;
  let inBracket = false;
  let escaped = false;
  for (const ch of code) {
    if (escaped) {
      cur += ch;
      escaped = false;
      continue;
    }
    if (ch === "\\") {
      cur += ch;
      escaped = true;
      continue;
    }
    if (ch === '"') {
      inQuote = !inQuote;
      cur += ch;
      continue;
    }
    if (!inQuote && ch === "[") inBracket = true;
    if (!inQuote && ch === "]") inBracket = false;
    if (ch === ";" && !inQuote && !inBracket) {
      out.push(cur);
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}

function parseSection(raw: string): Section {
  let code = "";
  let color: string | null = null;
  let condition: Section["condition"] = null;
  let i = 0;
  while (i < raw.length) {
    if (raw[i] === "[") {
      const end = raw.indexOf("]", i);
      if (end === -1) {
        code += raw.slice(i);
        break;
      }
      const inner = raw.slice(i + 1, end);
      const cm = CONDITION_BRACKET.exec(`[${inner}]`);
      if (cm) {
        condition = { op: cm[1] as ConditionOp, value: Number(cm[2]) };
      } else if (COLOR_BRACKET.test(`[${inner}]`)) {
        color = colorCodeToCss(inner);
      } else if (/^\$[^-]*-/.test(inner)) {
        // [$display-locale] 货币/区域标记（[$-409] / [$$-en-US] / [$¥-804]）：
        // display 部分按字面量显示、locale 部分不显示（ECMA-376 §18.8.30）。
        // 转成引号串后，数字段的边字面量回注与日期段的字面量扫描都按既有
        // 路径处理（此前整段原样保留，预览会显示出 "[$-409]" 字样）。
        const display = inner.slice(1, inner.indexOf("-", 1));
        code += display ? `"${display}"` : "";
      } else {
        // 未知括号（如某些方言的本地化关键字）：原样保留交引擎
        code += raw.slice(i, end + 1);
      }
      i = end + 1;
      continue;
    }
    // 引号串整体透传（保留给字面量抽取）
    if (raw[i] === '"') {
      const end = raw.indexOf('"', i + 1);
      const seg = end === -1 ? raw.slice(i) : raw.slice(i, end + 1);
      code += seg;
      i += seg.length;
      continue;
    }
    code += raw[i];
    i++;
  }
  return { raw, code, color, condition };
}

/** 数字记号字符（数字段的中段边界）。@ 为文本占位记号。 */
const NUM_TOKEN_CHARS = new Set("0#?.,%@");

/**
 * 抽取段码两侧"可安全回注"的字面量。
 * 规则：从两端向内扫描——引号串与 \ 转义原样保留；_ 与 * 连同其后一个字符
 * 丢弃（宽度填充，Excel 视觉上近似为无）；其余非记号字符（货币符号、-、
 * (、)、空格等）按字面量保留；遇到首个数字/文本记号即停。
 * 中段（含交错字面量，如电话号码 000-0000）的字面量丢失为已知降级。
 */
function extractEdgeLiterals(code: string): {
  prefix: string;
  suffix: string;
  core: string;
} {
  // 前缀
  let prefix = "";
  let i = 0;
  while (i < code.length) {
    const ch = code[i];
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const seg = end === -1 ? code.slice(i + 1) : code.slice(i + 1, end);
      prefix += seg;
      i += seg.length + 2;
      continue;
    }
    if (ch === "\\") {
      if (code[i + 1] !== undefined) prefix += code[i + 1];
      i += 2;
      continue;
    }
    if (ch === "_" || ch === "*") {
      i += 2; // 填充记号吞掉后一个字符
      continue;
    }
    if (NUM_TOKEN_CHARS.has(ch)) break;
    prefix += ch;
    i++;
  }
  // 后缀（对称，从尾向头）
  let suffix = "";
  let j = code.length - 1;
  while (j >= i) {
    const ch = code[j];
    if (ch === '"') {
      const start = code.lastIndexOf('"', j - 1);
      if (start < i) {
        suffix = code.slice(i, j + 1) + suffix;
        j = i - 1;
        break;
      }
      suffix = code.slice(start + 1, j) + suffix;
      j = start - 1;
      continue;
    }
    if (
      j > 0 &&
      (code[j - 1] === "\\" || code[j - 1] === "_" || code[j - 1] === "*")
    ) {
      if (code[j - 1] === "\\") suffix = ch + suffix;
      j -= 2;
      continue;
    }
    if (NUM_TOKEN_CHARS.has(ch)) break;
    suffix = ch + suffix;
    j--;
  }
  const core = code.slice(i, j + 1);
  return { prefix, suffix, core };
}

/**
 * 日期/时间段自实现（引擎缺陷倒逼，实测 2026-09）：
 *  - 引擎把 mm 一律当月份（h:mm 的分钟邻接规则未实现：0.5 + "h:mm" →
 *    "12:12"，mm=12 月；正确为 "12:00"）；
 *  - dddd（星期名）渲染成日号（45678.5 + "dddd" → "21"）。
 * 正常项保留在自实现里对齐：月份名（mmm）、AM/PM、yy/yyyy、d/dd、ss。
 * 时区口径：全程 UTC 分量（SheetJS 官方建议，规避 DST 重叠/空隙）。
 */
const MONTH_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];
const MONTH_LONG = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];
const DAY_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const DAY_LONG = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
];

interface DateTokens {
  tok: string;
  isMinute: boolean;
}

/** 扫码为记号流（含 m 的分钟/月份邻接消歧：前邻 h 或后邻 s 为分钟）。 */
function scanDateTokens(code: string): (DateTokens | string)[] {
  const out: (DateTokens | string)[] = [];
  let i = 0;
  while (i < code.length) {
    const ch = code[i];
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const seg = end === -1 ? code.slice(i + 1) : code.slice(i + 1, end);
      out.push(seg);
      i += seg.length + 2;
      continue;
    }
    if (ch === "\\") {
      if (code[i + 1] !== undefined) out.push(code[i + 1]);
      i += 2;
      continue;
    }
    const lower = ch.toLowerCase();
    if ("ymdhs".includes(lower) && /[a-z]/i.test(ch)) {
      let run = ch;
      while (
        i + run.length < code.length &&
        code[i + run.length].toLowerCase() === lower
      ) {
        run += code[i + run.length];
      }
      // 邻接消歧：直接前一个记号是 h，或紧跟（允许一个分隔符）s 记号 → 分钟。
      // 比较必须大小写不敏感（Excel 格式码记号本身大小写等价，第三方生成器
      // 常写大写 HH:MM:SS）——此前 tok[0] === "h" 按原字符比较，大写 H:MM 的
      // MM 被当月份渲染成"6:12"（12 月），与下方 formatDateSection 的
      // toLowerCase 口径不一致（实测踩坑）
      const prevTok = [...out].reverse().find((t) => typeof t !== "string");
      const nextIsS = /^[\s:.\/]?s/i.test(code.slice(i + run.length));
      const isMinute =
        lower === "m" && (prevTok?.tok[0].toLowerCase() === "h" || nextIsS);
      out.push({ tok: run, isMinute });
      i += run.length;
      continue;
    }
    if (code.slice(i).toUpperCase().startsWith("AM/PM")) {
      out.push({ tok: "AM/PM", isMinute: false });
      i += 5;
      continue;
    }
    if (code.slice(i).toUpperCase().startsWith("A/P")) {
      out.push({ tok: "A/P", isMinute: false });
      i += 3;
      continue;
    }
    // 普通字符（含分隔符与 .0 小数秒后缀）
    out.push(ch);
    i++;
  }
  return out;
}

/**
 * 自实现日期/时间格式化。serial 为 1900 系（date1904 由调用方 +1462 平移）。
 * 负值日期 Excel 显示 ####，此处对齐。
 */
export function formatDateSection(
  serial: number,
  code: string,
  serialToDate: (s: number) => Date | null,
): string {
  if (serial < 0) return "#######";
  const d = serialToDate(serial);
  if (!d || Number.isNaN(d.getTime())) return "#######";

  const days = Math.floor(serial);
  const frac = serial - days;
  const totalSec = Math.round(frac * SECS_PER_DAY);
  // 四舍五入到 86400（当天最后半秒内）→ 时间归 0 点且日期进到次日（Excel
  // 显示次日 0:00；此前日期不联动，1.9999999 显示 "1/1 0:00"）。
  // 进位日期不能靠 d + 1 天推：引擎 serialToDate 在幽灵日邻域自身会进位
  //（60.99999… 直接返回 3/1 0 点）而普通区间不进位（1.9999999 返回
  // 23:59:59.992），再 +1 天会重复进位——改用取整后的序列号回询引擎
  //（+0.5 锚在当天正午，日期分量稳定）。
  const rollDay = totalSec >= SECS_PER_DAY;
  const serialDay = days + (rollDay ? 1 : 0);
  const dd = rollDay ? (serialToDate(serialDay + 0.5) ?? d) : d;
  // 1900 幽灵日：serialDay 60 是 Excel 显示的 1900-02-29（真实不存在，继承
  // 自 Lotus 1-2-3）。引擎 serialToDate 把它折叠成 2/28（core 实测），这里按
  // Excel 显示口径回补一天（覆盖取整进位与未进位两种落入形态）。星期几无需
  // 调整——2/28/1900 与 Excel 的 2/29/1900 同为周三（Excel 的前置闰年偏差
  // 在该日恰好抵消）。
  const dayOfMonth = serialDay === 60 ? 29 : dd.getUTCDate();
  const secsOfDay = rollDay ? totalSec - SECS_PER_DAY : totalSec;
  // 浮点秒仅小数秒（ss.0）显示路径使用；整秒路径保持 round 口径不变
  const totalSecFloat = frac * SECS_PER_DAY;
  const secsOfDayFloat =
    totalSecFloat >= SECS_PER_DAY
      ? totalSecFloat - SECS_PER_DAY
      : totalSecFloat;
  const hour24 = Math.floor(secsOfDay / 3600);
  const minute = Math.floor((secsOfDay % 3600) / 60);
  const second = secsOfDay % 60;

  const tokens = scanDateTokens(code);
  const hasAmPm =
    tokens.some(
      (t) => typeof t !== "string" && t.tok.toUpperCase() === "AM/PM",
    ) ||
    tokens.some((t) => typeof t !== "string" && t.tok.toUpperCase() === "A/P");
  const pad = (n: number, w = 2) => String(n).padStart(w, "0");

  let out = "";
  for (let ti = 0; ti < tokens.length; ti++) {
    const t = tokens[ti];
    if (typeof t === "string") {
      out += t;
      continue;
    }
    const { tok, isMinute } = t;
    const head = tok[0].toLowerCase();
    const len = tok.length;
    if (head === "y") {
      out +=
        len >= 3 ? String(dd.getUTCFullYear()) : pad(dd.getUTCFullYear() % 100);
    } else if (head === "m") {
      if (isMinute) {
        out += len >= 2 ? pad(minute) : String(minute);
      } else {
        const m = dd.getUTCMonth();
        out +=
          len >= 4
            ? MONTH_LONG[m]
            : len === 3
              ? MONTH_SHORT[m]
              : len === 2
                ? pad(m + 1)
                : String(m + 1);
      }
    } else if (head === "d") {
      const wd = dd.getUTCDay();
      out +=
        len >= 4
          ? DAY_LONG[wd]
          : len === 3
            ? DAY_SHORT[wd]
            : len === 2
              ? pad(dayOfMonth)
              : String(dayOfMonth);
    } else if (head === "h") {
      const h = hasAmPm ? hour24 % 12 || 12 : hour24;
      out += len >= 2 ? pad(h) : String(h);
    } else if (head === "s") {
      // 小数秒（ss.0，内置 47 形态）：紧随的 ".0+" 字面量决定小数位，按
      // 显示精度四舍五入；这些字面量记号在此消费，避免 ".0" 被原样拼接。
      // 进位边界（59.96s → "60.0" 而分钟不进位）是已知近似：仅在真实秒数
      // 落入分钟翻转前半个显示精度内才出现，浮点尾差同量级，不值得为它
      // 引入分钟级联重算。
      let fracDigits = 0;
      const j = ti + 1;
      if (tokens[j] === ".") {
        let k = j + 1;
        while (tokens[k] === "0") k++;
        if (k > j + 1) fracDigits = k - (j + 1);
      }
      if (fracDigits > 0) {
        const secInMinute = secsOfDayFloat % 60;
        out += secInMinute
          .toFixed(fracDigits)
          .padStart(len + 1 + fracDigits, "0");
        ti = j + fracDigits; // 跳过 "." 与全部 "0" 字面量
        continue;
      }
      out += len >= 2 ? pad(second) : String(second);
    } else if (tok.toUpperCase() === "AM/PM") {
      out += hour24 < 12 ? "AM" : "PM";
    } else if (tok.toUpperCase() === "A/P") {
      out += hour24 < 12 ? "A" : "P";
    } else {
      out += tok;
    }
  }
  return out;
}
function isDateSection(code: string): boolean {
  let inQuote = false;
  let escaped = false;
  for (const ch of code) {
    if (escaped) {
      escaped = false;
      continue;
    }
    if (ch === "\\") {
      escaped = true;
      continue;
    }
    if (ch === '"') {
      inQuote = !inQuote;
      continue;
    }
    if (inQuote) continue;
    if ("ymdhs".includes(ch.toLowerCase()) && /[a-z]/i.test(ch)) return true;
  }
  return false;
}

/** 累计时长段：含 [h]/[m]/[s] 括号记号。 */
const ELAPSED_TOKEN = /\[(h+|m+|s+)\]/i;

/** 单位换算：Excel 一天 = 86400 秒。 */
const SECS_PER_DAY = 86_400;

/** 尾部秒记号（s+ 可带 .0+ 小数）的替换文本。 */
function elapsedSecondsText(sec: number, sm: RegExpExecArray): string {
  const frac = sm[1] ? sm[1].length - 1 : 0;
  const digits = sm[0].length - (sm[1]?.length ?? 0);
  if (frac > 0) {
    // 小数秒时 toFixed 输出含 "." 与尾数，补零只作用在整数位——padStart
    // 打在整串上会在个位数秒丢补零（[h]:mm:ss.0 → "6:00:0.0"，应为
    // "6:00:00.0"；日期段 formatDateSection 按 len+1+frac 总宽补齐无此问题）
    const [ip, fp] = sec.toFixed(frac).split(".");
    return ip.padStart(digits, "0") + "." + fp;
  }
  return String(Math.floor(sec)).padStart(digits, "0");
}

function formatElapsed(absSerial: number, code: string): string | null {
  // 引号与转义只是字面量的界定符、不属于显示文本（日期段由 scanDateTokens
  // 剥离）。字面量内容必须原样保留、且不参与记号扫描：直接删掉引号会把
  // [h]"h"mm"m"ss"s" 拼成 [h]hmmmsss，被 m+/s+ 正则吞成 "5h004059"。
  // 这里把引号串与 \x 转义抽成占位符（\0索引\0，不含 h/m/s），最后再回填。
  const literals: string[] = [];
  const hold = (text: string): string => {
    literals.push(text);
    return `\u0000${literals.length - 1}\u0000`;
  };
  let pat = "";
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      pat += hold(code.slice(i + 1, end === -1 ? code.length : end));
      i = end === -1 ? code.length : end;
      continue;
    }
    if (ch === "\\") {
      // 转义符：后一个字符按字面量显示（与日期段 scanDateTokens 同口径）
      if (code[i + 1] !== undefined) {
        pat += hold(code[i + 1]);
        i++;
      }
      continue;
    }
    pat += ch;
  }
  const restore = (s: string): string =>
    s.replace(/\u0000(\d+)\u0000/g, (_m, i: string) => literals[Number(i)]);
  const m = ELAPSED_TOKEN.exec(pat);
  if (!m) return null;
  const kind = m[1][0].toLowerCase();
  const width = m[1].length;
  const total = Math.abs(absSerial) * SECS_PER_DAY;

  if (kind === "s") {
    // [s] / [ss].0：整段即秒数，小数位由紧随的 .0+ 决定。该字面量记号必须
    // 随替换一并消费——此前只换掉 [ss] 本体，".0" 残留成 "21600.0.0"
    const after = pat.slice((m.index ?? 0) + m[0].length);
    const fm = /\.0+/.exec(after);
    const frac = fm ? fm[0].length - 1 : 0;
    const txt = frac > 0 ? total.toFixed(frac) : String(Math.floor(total));
    return restore(
      pat.slice(0, m.index) + txt + (fm ? after.slice(fm[0].length) : after),
    );
  }

  let head: number;
  let restSeconds = 0; // 括号记号取整后剩余的秒数，分配给尾部 m+/s+ 记号
  if (kind === "h") {
    head = Math.floor(total / 3600);
    restSeconds = total - head * 3600;
  } else {
    head = Math.floor(total / 60);
    restSeconds = total - head * 60;
  }

  const before = pat.slice(0, m.index);
  let tail = pat.slice((m.index ?? 0) + m[0].length);
  if (kind === "h") {
    const restMin = Math.floor(restSeconds / 60);
    const restSec = restSeconds - restMin * 60;
    tail = tail.replace(/m+/, (t) => String(restMin).padStart(t.length, "0"));
    const sm = /s+(\.0+)?/.exec(tail);
    if (sm) {
      tail = tail.replace(sm[0], elapsedSecondsText(restSec, sm));
    }
  } else {
    // [m]:ss——尾部只余秒
    const sm = /s+(\.0+)?/.exec(tail);
    if (sm) {
      tail = tail.replace(sm[0], elapsedSecondsText(restSeconds, sm));
    }
  }
  return restore(before + String(head).padStart(width, "0") + tail);
}

/**
 * 科学计数形态对齐 Excel：剥尾数零、指数至少两位（5e-8 → "5E-08"）。
 * 覆盖 [1e-10, 1e-7) 区间——toPrecision(15) 在该区间已切指数形态，此前
 * 该分支只 uppercase 不剥尾零，输出 "5.00000000000000E-8"。
 */
function sciNotation(s: string): string {
  return s
    .replace(/(\.\d*?)0+e/i, "$1e")
    .replace(/\.e/i, "e")
    .toUpperCase()
    .replace(
      /E([+-])(\d+)$/,
      (_m, sign: string, digits: string) =>
        `E${sign}${digits.padStart(2, "0")}`,
    );
}

function plainNotation(s: string): string {
  return s.includes(".") ? s.replace(/0+$/, "").replace(/\.$/, "") : s;
}

/** Excel General：15 位有效数字、去尾零；超宽量级用科学计数（近似）。 */
export function generalNumber(n: number): string {
  if (n === 0) return "0";
  if (!Number.isFinite(n)) return String(n);
  const abs = Math.abs(n);
  if (abs >= 1e15 || abs < 1e-10) {
    return sciNotation(n.toPrecision(6));
  }
  const s = n.toPrecision(15);
  // toPrecision(15) 在 [1e-10, 1e-7) 会切指数形态（1e10..1e14 之间理论上
  // 不会，走该分支属防御性兜底），两种形态统一经 sciNotation 归一
  return s.includes("e") || s.includes("E") ? sciNotation(s) : plainNotation(s);
}

/** 引擎调用的统一防御包装：任何异常降级为 String(v)。 */
function engineText(
  v: number | string | boolean,
  code: string,
): { text: string; color?: string } {
  try {
    const r = formatCellRich(v, code);
    if (r && typeof r.text === "string") {
      return typeof r.color === "string"
        ? { text: r.text, color: r.color }
        : { text: r.text };
    }
  } catch {
    /* 降级 */
  }
  return { text: String(v) };
}

/**
 * 格式化一个单元格的显示值。
 * @param type 归一化单元格类型
 * @param value 原始值（字符串形态，数字为字符串）
 * @param numFmtCode 已解析的格式码（builtin-table.resolveFormatCode 产物）
 * @param dateSystem 日期系统（1904 需平移序列号）
 */
export function formatCellValue(
  type: PreviewCellType,
  value: string | null,
  numFmtCode: string,
  dateSystem: "date1900" | "date1904",
): FormattedValue {
  // 空值（含公式无缓存值的单元格）一律空串
  if (value === null || value === undefined || value === "") {
    return { text: "" };
  }
  if (type === "boolean") {
    return {
      text: value === "1" || value.toLowerCase() === "true" ? "TRUE" : "FALSE",
    };
  }
  if (type === "error") {
    return { text: value };
  }

  const code = numFmtCode || "General";
  if (code === "General") {
    if (type === "number") {
      const n = Number(value);
      return { text: Number.isFinite(n) ? generalNumber(n) : value };
    }
    return { text: value };
  }

  const sections = splitSections(code).map(parseSection);
  const isNumericValue =
    type === "number" ||
    (type !== "string" &&
      type !== "formulaStr" &&
      Number.isFinite(Number(value)));

  if (!isNumericValue) {
    // 文本：找含 @ 的段（通常是第 4 段）；没有则原样显示
    const at = sections.find((s) => s.code.includes("@"));
    if (!at) return { text: value };
    const { prefix, suffix, core } = extractEdgeLiterals(at.code);
    const inner = engineText(value, core || "@");
    return {
      text: prefix + inner.text + suffix,
      color: at.color ?? inner.color,
    };
  }

  const n = Number(value);

  // 段选择：条件段按序匹配；否则按 Excel 正/负/零规则
  // （2 段 = 正;负——零走正段；3 段 = 正;负;零；单段全适用）
  let chosen: Section | undefined;
  for (const s of sections) {
    if (!s.condition) continue;
    const { op, value: cv } = s.condition;
    const ok =
      (op === ">=" && n >= cv) ||
      (op === ">" && n > cv) ||
      (op === "<=" && n <= cv) ||
      (op === "<" && n < cv) ||
      (op === "=" && n === cv) ||
      (op === "<>" && n !== cv);
    if (ok) {
      chosen = s;
      break;
    }
  }
  // 有条件段但无一命中：Excel 用"第一个无条件段"作 else（[<0]A;[>0]B;C 命中
  // C），而不是 section[0]——此前 5 配 [<0]"neg";[>1000]"big";0 会输出 "neg5"。
  if (!chosen && sections.some((s) => s.condition)) {
    chosen = sections.find((s) => !s.condition);
  }
  let signFromSingle = false;
  if (!chosen) {
    if (sections.length === 1) {
      chosen = sections[0];
      signFromSingle = n < 0; // 单段格式负值默认带负号（Excel 语义）
    } else if (n > 0) {
      chosen = sections[0];
    } else if (n < 0) {
      chosen = sections[1] ?? sections[0];
    } else {
      chosen = sections.length >= 3 ? sections[2] : sections[0];
    }
  }

  // 累计时长（含 [h]/[m]/[s]）：自实现。无括号 mm:ss 不在此拦截——它是
  // 当日时间的"分:秒"（Excel 语义；累计分钟是带括号 [m]:ss 的语义），
  // 随下方日期段走 formatDateSection（分钟邻接消歧已覆盖 mm:ss 形态）。
  if (ELAPSED_TOKEN.test(chosen.code)) {
    // 负值时长 Excel 显示 ####（与下方日期段 formatDateSection 口径一致），
    // 此前按绝对值渲染会把 -0.25 天显示成 "6:00:00"
    if (n < 0) return { text: "#######", color: chosen.color ?? undefined };
    const t = formatElapsed(n, chosen.code);
    if (t !== null) return { text: t, color: chosen.color ?? undefined };
  }

  // 零值走零段：实测引擎对零段字面量无损，整段直交（非日期形态）
  if (
    n === 0 &&
    sections.length >= 3 &&
    chosen === sections[2] &&
    !isDateSection(chosen.code) &&
    !ELAPSED_TOKEN.test(chosen.code)
  ) {
    const r = engineText(0, chosen.code);
    return { text: r.text, color: chosen.color ?? r.color };
  }

  // 日期/时间段：自实现（引擎 mm 恒为月份、dddd 渲染为日号，实测缺陷）；
  // 1904 系统平移 1462 天（serialToDate 固定 1900 epoch，实测签名无参数）
  if (isDateSection(chosen.code)) {
    const serial = dateSystem === "date1904" ? n + 1462 : n;
    const text = formatDateSection(serial, chosen.code, serialToDate);
    return { text, color: chosen.color ?? undefined };
  }

  // 数字段：拆边字面量，|值| 交引擎，再拼回；单段负值补前导负号
  const { prefix, suffix, core } = extractEdgeLiterals(chosen.code);
  const r = engineText(Math.abs(n), core || "General");
  const sign = signFromSingle ? "-" : "";
  return {
    text: sign + prefix + r.text + suffix,
    color: chosen.color ?? r.color,
  };
}
