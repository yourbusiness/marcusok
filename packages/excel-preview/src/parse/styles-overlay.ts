/**
 * 样式覆盖层：自解 styles.xml + theme1.xml，承担两件引擎读取侧做不到的事
 * （均有实测依据，2026-09）：
 *
 *  1. 主题色/indexed 颜色找回：modern-xlsx 把 <color theme="…"/>、
 *     <color indexed="…"/> 一律解析成 null——真实 Excel 文件的默认字体色
 *     （theme 1）与大量主题色填充（"良好/中立"样式、accent 系列表头）会
 *     整体丢失。
 *
 *  2. fonts/fills/borders 集合按 styles.xml 文档序完整重建：引擎解析会跳过
 *     自闭合元素（<border/>、<font/>、<fill/>——真实 Excel 产物的 border#0
 *     几乎总是自闭合 <border/>），数组因此比文件少一项，cellXfs 的
 *     fontId/fillId/borderId（文件下标）整体错位：borderId=0 的默认单元格
 *     会拿到文件中 border#1 的样式、末位引用越界丢失。本层输出含自闭合
 *     空项的完整列表，与 cellXfs 的 id 语义严格 1:1（模型构建优先用它，
 *     见 model.ts 的 buildStyles；引擎数组仅在覆盖层不可用时降级使用）。
 *
 * 只解压 styles/theme 两个小部件（styles.xml 通常几 KB，sheetData 大部件
 * 不解）；任何异常吞掉、返回降级形态（调用方回退引擎数组）。
 */
import { unzipSync } from "fflate";
import {
  applyTint,
  argbToCss,
  indexedColor,
  themeColor,
  type ThemePalette,
} from "../color";

/** styles.xml / theme1.xml 中的原始颜色规格。 */
export interface RawColorSpec {
  kind: "rgb" | "theme" | "indexed" | "auto";
  /** rgb：8 位 ARGB 或 6 位 RGB。 */
  rgb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

/** font 元素的完整解析（含自闭合空项；文档序与 cellXfs 的 fontId 1:1）。 */
export interface OverlayFont {
  name: string | null;
  /** pt；null = 缺省（渲染层默认 11）。 */
  size: number | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: RawColorSpec | null;
}

/** border 单边：style 缺省（<left/>）即无边。 */
export interface OverlayBorderSide {
  /** ECMA-376 边框样式名（thin/medium/...）；null = 该边无边框。 */
  style: string | null;
  color: RawColorSpec | null;
}

export interface OverlayBorder {
  left: OverlayBorderSide | null;
  right: OverlayBorderSide | null;
  top: OverlayBorderSide | null;
  bottom: OverlayBorderSide | null;
  diagonal: OverlayBorderSide | null;
}

/** fill 元素的完整解析（pattern 与 gradient 二选一；均无 = kind none）。 */
export interface OverlayFill {
  /** patternType 属性（"none"/"solid"/"gray125"/…）；null = 自闭合或无 patternFill。 */
  patternType: string | null;
  fgColor: RawColorSpec | null;
  bgColor: RawColorSpec | null;
  gradient: {
    degree: number;
    stops: { position: number; spec: RawColorSpec | null }[];
  } | null;
}

export interface StylesOverlay {
  /** 三个集合均按 styles.xml 文档序、含自闭合空项（与 cellXfs id 对齐）。 */
  fonts: OverlayFont[];
  fills: OverlayFill[];
  borders: OverlayBorder[];
}

/** 只解压需要的部件（filter 避免 inflate 巨大的 sheetData）。 */
function readParts(bytes: Uint8Array): { styles?: string; theme?: string } {
  const parts: { styles?: string; theme?: string } = {};
  // 局部 zip 目录损坏等异常一律吞掉——覆盖层是尽力而为的增强，失败即降级
  try {
    const unzipped = unzipSync(bytes, {
      filter: (f) =>
        f.name === "xl/styles.xml" || f.name === "xl/theme/theme1.xml",
    });
    const decoder = new TextDecoder();
    if (unzipped["xl/styles.xml"])
      parts.styles = decoder.decode(unzipped["xl/styles.xml"]);
    if (unzipped["xl/theme/theme1.xml"])
      parts.theme = decoder.decode(unzipped["xl/theme/theme1.xml"]);
  } catch {
    /* 降级：无覆盖层 */
  }
  return parts;
}

/** XML 属性表解析：`<color rgb="FF0000FF" theme="1"/>` → {rgb, theme}。 */
function parseAttrs(tag: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const re = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(tag)) !== null) attrs[m[1]] = decodeXmlEntities(m[2]);
  return attrs;
}

/** 最小 XML 实体解码（颜色值不含实体，formatCode 可能含 &quot;）。 */
function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, "&");
}

/** 属性文本（`rgb="…" theme="…" tint="…"`）→ 颜色规格。 */
function specFromAttrs(attrText: string): RawColorSpec | null {
  const attrs = parseAttrs(attrText);
  if (attrs.rgb !== undefined) return { kind: "rgb", rgb: attrs.rgb };
  if (attrs.theme !== undefined) {
    return {
      kind: "theme",
      theme: Number(attrs.theme),
      tint: attrs.tint !== undefined ? Number(attrs.tint) : undefined,
    };
  }
  if (attrs.indexed !== undefined) {
    return { kind: "indexed", indexed: Number(attrs.indexed) };
  }
  if (attrs.auto !== undefined) return { kind: "auto" };
  return null;
}

/** 从元素内容里提取 <color .../> 规格（字体/边框/渐变 stop 用）。 */
function parseColorSpec(content: string): RawColorSpec | null {
  const m = /<color\s+([^>]*?)\/>/.exec(content);
  return m ? specFromAttrs(m[1]) : null;
}

/** 解析 theme1.xml 的 clrScheme → 调色板（srgbClr val / sysClr lastClr）。 */
export function parseThemePalette(themeXml: string): ThemePalette {
  const scheme = /<a:clrScheme[\s\S]*?<\/a:clrScheme>/.exec(themeXml);
  if (!scheme) return {};
  const palette: ThemePalette = {};
  const entry =
    /<a:(dk1|lt1|dk2|lt2|accent1|accent2|accent3|accent4|accent5|accent6|hlink|folHlink)>([\s\S]*?)<\/a:\1>/g;
  let m: RegExpExecArray | null;
  while ((m = entry.exec(scheme[0])) !== null) {
    const srgb = /<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(m[2]);
    const sys = /<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/.exec(m[2]);
    const v = srgb?.[1] ?? sys?.[1];
    if (v) palette[m[1] as keyof ThemePalette] = v;
  }
  return palette;
}

function splitElements(xml: string, tag: string): string[] {
  // 两种形态都必须产出列表项：配对 <tag …>…</tag> 与自闭合 <tag …/>（真实
  // Excel 产物的常态，如默认无边框 <border/>）。自闭合产出空内容项——输出
  // 下标因此与 styles.xml 文档序严格 1:1，cellXfs 的 id 引用才有对齐基准
  // （引擎解析会跳过自闭合元素导致其数组错位，见文件头注释第 2 点）。
  const re = new RegExp(`<${tag}\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${tag}>)`, "g");
  const out: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) out.push(m[2] ?? "");
  return out;
}

/** 布尔子元素（<b/>、<b val="1"/>）：<b val="0"/> 是显式否定，非缺省。 */
function flagElem(content: string, tag: string): boolean {
  const re = new RegExp(`<${tag}\\b([^>]*?)(?:/>|>)`);
  const m = re.exec(content);
  if (!m) return false;
  const val = parseAttrs(m[1]).val;
  return val === undefined || val === "1" || val === "true";
}

function parseFontElem(content: string): OverlayFont {
  const numVal = (tag: string): number | null => {
    const m = new RegExp(`<${tag}\\b([^>]*?)(?:/>|>)`).exec(content);
    const v = m ? parseAttrs(m[1]).val : undefined;
    const n = v !== undefined ? Number(v) : NaN;
    return Number.isFinite(n) && n >= 0 ? n : null;
  };
  const name = /<name\b[^>]*?val="([^"]*)"/.exec(content)?.[1] ?? null;
  return {
    name: name || null,
    size: numVal("sz"),
    bold: flagElem(content, "b"),
    italic: flagElem(content, "i"),
    underline: flagElem(content, "u"),
    strike: flagElem(content, "strike"),
    color: parseColorSpec(content),
  };
}

function parseBorderElem(content: string): OverlayBorder {
  const side = (tag: string): OverlayBorderSide | null => {
    // <left/>（style 缺省）= 该边无边框，返回 null 与引擎形态一致
    const m = new RegExp(
      `<${tag}\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${tag}>)`,
    ).exec(content);
    if (!m) return null;
    const style = parseAttrs(m[1]).style;
    if (!style) return null;
    return {
      style,
      color: m[2] !== undefined ? parseColorSpec(m[2]) : null,
    };
  };
  return {
    left: side("left"),
    right: side("right"),
    top: side("top"),
    bottom: side("bottom"),
    diagonal: side("diagonal"),
  };
}

function parseFillElem(content: string): OverlayFill {
  const out: OverlayFill = {
    patternType: null,
    fgColor: null,
    bgColor: null,
    gradient: null,
  };
  const pat = /<patternFill\b([^>]*?)(?:\/>|>([\s\S]*?)<\/patternFill>)/.exec(
    content,
  );
  if (pat) {
    out.patternType = parseAttrs(pat[1]).patternType ?? "";
    const inner = pat[2] ?? "";
    // patternFill 用 <fgColor>/<bgColor> 元素名（非 <color>），单独提取
    const fg = /<fgColor\s+([^>]*?)\/>/.exec(inner);
    const bg = /<bgColor\s+([^>]*?)\/>/.exec(inner);
    out.fgColor = fg ? specFromAttrs(fg[1]) : null;
    out.bgColor = bg ? specFromAttrs(bg[1]) : null;
    return out;
  }
  const grad =
    /<gradientFill\b([^>]*?)(?:\/>|>([\s\S]*?)<\/gradientFill>)/.exec(content);
  if (grad) {
    const degree = Number(parseAttrs(grad[1]).degree ?? "90");
    const stops: { position: number; spec: RawColorSpec | null }[] = [];
    const stopRe = /<stop\s+([^>]*?)>([\s\S]*?)<\/stop>/g;
    let sm: RegExpExecArray | null;
    while ((sm = stopRe.exec(grad[2] ?? "")) !== null) {
      const pos = Number(parseAttrs(sm[1]).position ?? "0");
      stops.push({
        position: Number.isFinite(pos) ? pos : 0,
        spec: parseColorSpec(sm[2]),
      });
    }
    out.gradient = { degree: Number.isFinite(degree) ? degree : 90, stops };
  }
  return out;
}

/** 从 zip 字节构建覆盖层（styles.xml 缺失时返回 null，调用方降级）。 */
export function buildStylesOverlay(bytes: Uint8Array): {
  overlay: StylesOverlay | null;
  palette: ThemePalette;
} {
  const parts = readParts(bytes);
  const palette = parts.theme ? parseThemePalette(parts.theme) : {};
  if (!parts.styles) return { overlay: null, palette };
  const xml = parts.styles;

  const fontsSec = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(xml)?.[1] ?? "";
  const fillsSec = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(xml)?.[1] ?? "";
  const bordersSec =
    /<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(xml)?.[1] ?? "";

  const overlay: StylesOverlay = {
    fonts: splitElements(fontsSec, "font").map(parseFontElem),
    fills: splitElements(fillsSec, "fill").map(parseFillElem),
    borders: splitElements(bordersSec, "border").map(parseBorderElem),
  };

  return { overlay, palette };
}

/** 规格解析为最终 CSS 色；失败返回 null（调用方保持降级值）。 */
export function resolveColorSpec(
  spec: RawColorSpec | null | undefined,
  palette: ThemePalette,
): string | null {
  if (!spec) return null;
  switch (spec.kind) {
    case "rgb":
      return argbToCss(spec.rgb);
    case "theme": {
      if (spec.theme === undefined) return null;
      const base = themeColor(palette, spec.theme);
      if (!base) return null;
      if (spec.tint !== undefined && spec.tint !== 0) {
        return argbToCss(applyTint(base, spec.tint));
      }
      return argbToCss(base);
    }
    case "indexed":
      return spec.indexed === undefined
        ? null
        : argbToCss(indexedColor(spec.indexed));
    case "auto":
      return null;
  }
}
