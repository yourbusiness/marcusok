/**
 * 颜色解析：ARGB 规范化、主题色 tint、indexed 调色板。
 *
 * 存在的原因（实测 2026-09）：modern-xlsx 读取侧把 styles.xml 中的
 * theme/tint/indexed 颜色一律解析为 null（FontData/FillData.color 类型即
 * hexRGB | null，无 theme 字段），真实 Excel 文件的默认字体色（theme 1）与
 * 主题色填充会整体丢失。覆盖层（parse/styles-overlay.ts）自解 styles.xml
 * 拿回原始颜色规格，本模块负责把规格解析成最终 RGB。
 */

/**
 * 8 位 ARGB（如 "FF0563C1"）→ "#RRGGBB"。alpha 非 FF 时输出 rgba() 以保留
 * 半透明（罕见，但 CAD/条件格式导出里出现过）。非法输入返回 null。
 */
export function argbToCss(argb: string | null | undefined): string | null {
  if (!argb) return null;
  let s = String(argb).trim().toLowerCase();
  if (s.startsWith("#")) s = s.slice(1);
  if (!/^[0-9a-f]+$/.test(s)) return null;
  if (s.length === 8) {
    const a = parseInt(s.slice(0, 2), 16) / 255;
    const r = s.slice(2, 4);
    const g = s.slice(4, 6);
    const b = s.slice(6, 8);
    if (a >= 0.999) return `#${r}${g}${b}`;
    return `rgba(${parseInt(r, 16)}, ${parseInt(g, 16)}, ${parseInt(b, 16)}, ${Math.round(a * 100) / 100})`;
  }
  if (s.length === 6) return `#${s}`;
  return null;
}

/** xlsx <color theme="N"> 的取值 → 主题调色板键。0/1 与 2/3 相对 clrScheme
 * 文档序是交换的（Excel 默认字体色写 theme="1"，对应 dk1/windowText）。 */
const THEME_INDEX_TO_KEY = [
  "lt1",
  "dk1",
  "lt2",
  "dk2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

export type ThemePalette = Partial<
  Record<(typeof THEME_INDEX_TO_KEY)[number], string>
>;

export function themeColor(
  palette: ThemePalette,
  themeIndex: number,
): string | null {
  const key = THEME_INDEX_TO_KEY[themeIndex];
  const v = key ? palette[key] : undefined;
  return v ? v : null;
}

/**
 * ECMA-376 §18.8.19 tint：对颜色的 HSL 亮度施加 tint（-1..1）。
 * tint < 0：L' = L × (1 + tint)（变暗）；tint > 0：L' = L × (1 - tint) + tint
 * （变亮，归一化形式与 MS-OI29500 的 HLSMAX 公式等价）。Excel 的 HLS 用
 * 240 刻度，比值公式与 0..1 归一化在数学上同构。
 */
export function applyTint(hex6: string, tint: number): string {
  const h = hex6.replace("#", "");
  const r = parseInt(h.slice(0, 2), 16) / 255;
  const g = parseInt(h.slice(2, 4), 16) / 255;
  const b = parseInt(h.slice(4, 6), 16) / 255;
  const [hh, ss, ll] = rgbToHsl(r, g, b);
  // tint > 0 分支的规范文本为 L' = L×(1-tint) + (HLSMAX - HLSMAX×(1-tint))，
  // 归一化（HLSMAX=1）后即 L×(1-tint) + tint。
  const l = tint < 0 ? ll * (1 + tint) : ll * (1 - tint) + tint;
  const [r2, g2, b2] = hslToRgb(hh, ss, Math.min(1, Math.max(0, l)));
  const to2 = (x: number) =>
    Math.round(Math.min(1, Math.max(0, x)) * 255)
      .toString(16)
      .padStart(2, "0");
  return `${to2(r2)}${to2(g2)}${to2(b2)}`;
}

/** sRGB → HSL（各通道 0..1）。 */
export function rgbToHsl(
  r: number,
  g: number,
  b: number,
): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
  else if (max === g) h = ((b - r) / d + 2) / 6;
  else h = ((r - g) / d + 4) / 6;
  return [h, s, l];
}

/** HSL → sRGB（各通道 0..1）。 */
export function hslToRgb(
  h: number,
  s: number,
  l: number,
): [number, number, number] {
  if (s === 0) return [l, l, l];
  const hue2rgb = (p: number, q: number, t: number) => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue2rgb(p, q, h + 1 / 3), hue2rgb(p, q, h), hue2rgb(p, q, h - 1 / 3)];
}

/**
 * ECMA-376 legacy indexed 调色板（§18.8.27 附表，en-US Excel 实际值）。
 * 0-63 为固定表；64/65 为系统前景/背景——取黑/白常见近似值。
 */
const INDEXED_COLORS: readonly string[] = [
  "000000",
  "ffffff",
  "ff0000",
  "00ff00",
  "0000ff",
  "ffff00",
  "ff00ff",
  "00ffff",
  "000000",
  "ffffff",
  "ff0000",
  "00ff00",
  "0000ff",
  "ffff00",
  "ff00ff",
  "00ffff",
  "800000",
  "008000",
  "000080",
  "808000",
  "800080",
  "008080",
  "c0c0c0",
  "808080",
  "9999ff",
  "993366",
  "ffffcc",
  "ccffff",
  "660066",
  "ff8080",
  "0066cc",
  "ccccff",
  "000080",
  "ff00ff",
  "ffff00",
  "00ffff",
  "800080",
  "800000",
  "008080",
  "0000ff",
  "00ccff",
  "ccffff",
  "ccffcc",
  "ffff99",
  "99ccff",
  "ff99cc",
  "cc99ff",
  "ffcc99",
  "3366ff",
  "33cccc",
  "99cc00",
  // 51/52（ccff00/ffcc00）此前缺失，53-62 整体前移两位、62/63 值错——
  // 对照 ECMA-376 §18.8.27 附表与 openpyxl/POI 标准实现逐一核对修正
  "ccff00",
  "ffcc00",
  "ff9900",
  "ff6600",
  "666699",
  "969696",
  "003366",
  "339966",
  "003300",
  "333300",
  "993300",
  "993366",
  "333399",
  "000000", // 64：system foreground
  "ffffff", // 65：system background
];

export function indexedColor(index: number): string | null {
  if (index < 0 || index >= INDEXED_COLORS.length) return null;
  return INDEXED_COLORS[index];
}

/** [Color n] / 命名色（[Red] 等）→ CSS 色（统一 # 前缀）。 */
export function colorCodeToCss(name: string): string | null {
  const m = /^color\s+(\d+)$/i.exec(name.trim());
  if (m) return argbToCss(indexedColor(Number(m[1])));
  const hex = NAMED_COLORS.get(name.trim().toLowerCase());
  return hex ? `#${hex}` : null;
}

/** numFmt 颜色段支持的命名色（en-US 名 → 6 位 RGB）。 */
const NAMED_COLORS = new Map<string, string>([
  ["black", "000000"],
  ["blue", "0000ff"],
  ["cyan", "00ffff"],
  ["green", "008000"],
  ["magenta", "ff00ff"],
  ["red", "ff0000"],
  ["white", "ffffff"],
  ["yellow", "ffff00"],
]);
