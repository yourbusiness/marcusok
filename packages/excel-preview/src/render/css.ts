/**
 * 样式表 → CSS 编译：每个 xf 生成一个 class（.xpv-xf-N），字体/填充/边框/
 * 对齐/缩进/换行/旋转全部拍平进该 class。单元格 DOM 只挂 class + 定位，
 * 保证虚拟滚动重建时样式计算零成本。
 *
 * 逐项映射依据（Excel 视觉约定）：
 *  - 字号用 pt 原值（CSS pt 与 Excel pt 同单位）
 *  - horizontal 缺省（General）按类型分流：数字右、文本左、布尔/错误居中——
 *    由渲染层附加 .xpv-num/.xpv-text/.xpv-bool 类实现，class 编译只处理
 *    显式对齐
 *  - 边框只画设了的边；相邻格边框同像素位重叠，后画覆盖先画，视觉正确
 *  - 网格线不在本表编译：渲染层以 .xpv-cell 的右/下 border 另行注入，
 *    且规则前置于 xf 类（同特异性级联后源序胜，数据边框覆盖网格线，
 *    见 renderer.ts 的 GRID_COLOR 说明）
 *
 * 选择器按实例作用域（scope）前缀：类名 .xpv-xf-N 是全局的、每个实例都会
 * 生成同名规则，同页挂载多个预览时后插入的 <style> 会按源序覆盖前者——
 * 两个预览的同一索引格子会互相串样式。scope 由 renderer 按实例唯一分配，
 * 规则编译成 `.xpv-s2 .xpv-xf-3{…}`；不传 scope（单元测试直接编译规则时）
 * 保持旧的无前缀输出。
 */
import type { PreviewStyles, PreviewXf } from "../types";

const BORDER_CSS: Record<string, string> = {
  thin: "1px solid",
  medium: "2px solid",
  thick: "3px solid",
  double: "3px double",
  dashed: "1px dashed",
  dotted: "1px dotted",
  hair: "1px solid",
  dashDot: "1px dashed",
  dashDotDot: "1px dashed",
  mediumDashDot: "2px dashed",
  mediumDashDotDot: "2px dashed",
  mediumDashed: "2px dashed",
  slantDashDot: "2px dashed",
};

function borderDecl(style: string, color: string | null): string {
  const base = BORDER_CSS[style] ?? "1px solid";
  return `${base} ${color ?? "#000000"}`;
}

/** 填充 → background-color 声明（none 省略）。渐变不在本函数输出：它与
 * 对角线边框同落 background-image，必须在 compileStylesheet 分层合成
 * （见 gradientLayer），否则两条独立声明按源序后者胜、渐变被冲掉。 */
function fillDecl(fill: PreviewStyles["fills"][number] | undefined): string {
  if (!fill) return "";
  switch (fill.kind) {
    case "solid":
      // fgColor null = 自动色：真实文件 solid 填充基本都有显式/主题色（覆盖
      // 层找回），落空到这里的极端罕见——保守按"无填充"渲染，避免整片黑
      return fill.fgColor ? `background-color:${fill.fgColor};` : "";
    case "pattern":
      // 图案填充近似：底色为主，图案色不还原（文档标注的近似项）
      return `background-color:${fill.bgColor ?? fill.fgColor ?? "#ffffff"};`;
    default:
      return "";
  }
}

/** 渐变填充的 background-image 层（null = 无）。 */
function gradientLayer(
  fill: PreviewStyles["fills"][number] | undefined,
): string | null {
  if (!fill || fill.kind !== "gradient") return null;
  const stops = fill.stops
    .map((s) => `${s.color} ${Math.round(s.position * 100)}%`)
    .join(", ");
  // xlsx degree：0=左→右、90=上→下（ECMA-376 CT_GradientFill）；CSS 线性
  // 渐变 0deg=向上、90deg=向右、180deg=向下，故 CSS 角度 = degree + 90
  const angle = ((fill.degree % 360) + 360) % 360;
  return `linear-gradient(${(angle + 90) % 360}deg, ${stops})`;
}

function fontDecl(font: PreviewStyles["fonts"][number] | undefined): string {
  if (!font) return "";
  let css = "";
  if (font.bold) css += "font-weight:700;";
  if (font.italic) css += "font-style:italic;";
  // 下划线与删除线同置必须合成一条 text-decoration（两条独立声明按源序
  // 后者胜，只剩 line-through）；此前修复合同时用了整串赋值，把已累加的
  // font-weight/font-style 一并冲掉——bold+下划线+删除线的字体丢失粗体
  if (font.underline && font.strike) {
    css += "text-decoration:underline line-through;";
  } else if (font.underline) {
    css += "text-decoration:underline;";
  } else if (font.strike) {
    css += "text-decoration:line-through;";
  }
  if (font.size) css += `font-size:${font.size}pt;`;
  // 字体名拼入 CSS 前剥掉双引号与反斜杠：双引号截断字符串、反斜杠转义
  // 闭合引号，都会让恶意 xlsx 的字体名破坏整段生成的样式规则（颜色等
  // 其余插值均经十六进制/命名表校验，字体名是唯一裸拼点）
  const family = font.name
    ? `"${font.name.replace(/["\\]/g, "")}", Calibri, "Segoe UI", system-ui, sans-serif`
    : `Calibri, "Segoe UI", system-ui, sans-serif`;
  css += `font-family:${family};`;
  if (font.color) css += `color:${font.color};`;
  return css;
}

function alignDecl(xf: PreviewXf): string {
  const a = xf.alignment;
  if (!a) return "";
  let css = "";
  // .xpv-cell 是 flex 容器（BASE_CSS）：文本是匿名 flex item，宽度收缩到
  // 内容宽（flex-grow 默认 0），text-align/vertical-align 对其无效果——
  // 主轴必须用 justify-content、交叉轴必须用 align-items 才能移动内容。
  // （此前写 text-align/vertical-align，显式水平/垂直对齐全静默失效，
  // 只有缺省类型分流路径的 .xpv-num/.xpv-text/.xpv-mid 是对的。）
  switch (a.horizontal) {
    case "left":
      css += "justify-content:flex-start;";
      break;
    case "center":
      css += "justify-content:center;";
      break;
    case "right":
      css += "justify-content:flex-end;";
      break;
    case "justify":
    case "distributed":
      // 两端对齐作用于换行后的行内文本（item 收缩到容器宽时 text-align
      // 有效），容器主轴保持左对齐即可
      css += "justify-content:flex-start;text-align:justify;";
      break;
    // general / fill / centerContinuous：general 走类型分流；fill/centerContinuous
    // 为罕见特性，按 general 近似
    default:
      break;
  }
  switch (a.vertical) {
    case "top":
      css += "align-items:flex-start;";
      break;
    case "center":
      css += "align-items:center;";
      break;
    case "bottom":
      css += "align-items:flex-end;";
      break;
    // justify/distributed 无 CSS 等价物，按居中近似（与前版一致）
    case "justify":
    case "distributed":
      css += "align-items:center;";
      break;
    default:
      break;
  }
  if (a.wrapText) css += "white-space:pre-wrap;word-break:break-word;";
  if (typeof a.indent === "number" && a.indent > 0) {
    // Excel 缩进单位 ≈ 3 个空格字符宽
    css += `padding-inline-start:${a.indent * 3}ch;`;
    if (a.horizontal === "right") {
      css += `padding-inline-end:${a.indent * 3}ch;padding-inline-start:0;`;
    }
  }
  return css;
}

/** 对角线边框的 background-image 层（null = 无）。CSS 无原生对角线，用
 * 线性渐变近似（细对角线）；方向由 PreviewBorder.diagonalUp 决定。 */
function diagonalLayer(
  border: PreviewStyles["borders"][number] | undefined,
): string | null {
  if (!border?.diagonal) return null;
  const c = border.diagonal.color ?? "#000000";
  const w =
    border.diagonal.style === "medium" || border.diagonal.style === "thick"
      ? "2px"
      : "1px";
  // 方向：diagonalUp = "/"（to top right）；缺省/Down = "\"（to bottom
  // right，Excel diagonalDown 是常见形态）
  const dir = border.diagonalUp ? "to top right" : "to bottom right";
  return `linear-gradient(${dir}, transparent calc(50% - ${w}), ${c} calc(50% - ${w}), ${c} calc(50% + ${w}), transparent calc(50% + ${w}))`;
}

function borderDeclFor(
  border: PreviewStyles["borders"][number] | undefined,
): string {
  if (!border) return "";
  // 只出四边声明；对角线在 diagonalLayer（与渐变填充合成 background-image）
  let css = "";
  if (border.left)
    css += `border-left:${borderDecl(border.left.style, border.left.color)};`;
  if (border.right)
    css += `border-right:${borderDecl(border.right.style, border.right.color)};`;
  if (border.top)
    css += `border-top:${borderDecl(border.top.style, border.top.color)};`;
  if (border.bottom)
    css += `border-bottom:${borderDecl(border.bottom.style, border.bottom.color)};`;
  return css;
}

/**
 * 编译整个样式表为一段 CSS（`.xpv-xf-0 … .xpv-xf-N`）。
 * @param scope 实例作用域类（renderer 分配的 `.xpv-sN`）；给出时每条规则
 *   编译为后代选择器，隔离同页多实例的同名类；缺省保持无前缀输出。
 */
export function compileStylesheet(
  styles: PreviewStyles,
  scope?: string,
): string {
  const parts: string[] = [];
  styles.xfs.forEach((xf, i) => {
    const font = styles.fonts[xf.fontId];
    const fill = styles.fills[xf.fillId];
    const border = styles.borders[xf.borderId];
    // 渐变填充与对角线边框都落 background-image：两条独立声明会按源序
    // 后者胜（先写的渐变被冲掉），必须合成一条逗号分隔的多层声明（对角线
    // 在上层）。此前渐变用 background 简写，同样会整体重置 image 层。
    const diag = diagonalLayer(border);
    const grad = gradientLayer(fill);
    const bgImage =
      diag || grad
        ? `background-image:${[diag, grad].filter(Boolean).join(",")};`
        : "";
    const css =
      fontDecl(font) +
      fillDecl(fill) +
      bgImage +
      borderDeclFor(border) +
      alignDecl(xf);
    if (css) {
      parts.push(
        scope ? `.${scope} .xpv-xf-${i}{${css}}` : `.xpv-xf-${i}{${css}}`,
      );
    }
  });
  return parts.join("\n");
}

/** 旋转样式（textRotation：1-90 逆时针，91-180 顺时针，255 竖排）。 */
export function rotationCss(
  rotation: number | null | undefined,
): string | null {
  if (rotation == null) return null;
  if (rotation === 0) return null;
  if (rotation === 255) return "writing-mode:vertical-rl;";
  if (rotation >= 1 && rotation <= 90) {
    return `transform:rotate(${-rotation}deg);transform-origin:left bottom;white-space:nowrap;`;
  }
  // 91-180：顺时针 (rotation-90) 度
  return `transform:rotate(${rotation - 90}deg);transform-origin:right bottom;white-space:nowrap;`;
}
