/**
 * 框架无关渲染器：DOM 四象限虚拟滚动。
 *
 * 结构（viewport 为滚动容器；content 撑出滚动尺寸；cells 在网格坐标系
 * 绝对定位，随容器自然滚动——滚动本身零 DOM 更新，只有窗口边缘的格子增删）：
 *
 *   root ─ tabs（可选）
 *        └ viewport
 *            ├ content ─ cellsLayer（主区格子，网格坐标）
 *            ├ frozenColLayer  （钉 X：translateX(+scrollLeft)）
 *            ├ frozenRowLayer  （钉 Y：translateY(+scrollTop)）
 *            ├ frozenCorner    （冻结角块，双向钉）
 *            ├ headerRow       （列标，钉 Y）
 *            ├ headerCol       （行号，钉 X）
 *            ├ colHeadFrozen   （冻结列的列标，双向钉）
 *            ├ rowHeadFrozen   （冻结行的行号，双向钉）
 *            └ headerCorner    （左上角，双向钉）
 *
 * 关键不变量：
 *  - 格子元素以 `r:c` 为 key 复用：同一 sheet 下同一格子的内容与几何恒定
 *    （模型不可变），创建时一次性写入定位/内容/内联样式，滚动只做集合差量
 *    （增删），绝不改已有元素——避免滚动帧内样式抖动与重排；
 *  - 覆盖层（含冻结层与表头）都是滚动容器 .xpv-viewport
 *    （position:relative/overflow:auto）的绝对定位子元素，**会随滚动一起
 *    位移**；因此各条带用 transform 抵消"需要钉住"的那个轴（合成器路径，
 *    不触发布局），而需要跟随滚动的轴不加 transform（随滚动自然位移）；
 *  - 合并单元格经 layout 的召回机制整体渲染（主单元格可能在窗口外）。
 */
import { formatCellValue } from "../numfmt/format";
import {
  buildLayout,
  computeVisibleRange,
  visibleMergeAnchors,
  columnIndexLabel,
  HEADER_COL_PX,
  HEADER_ROW_PX,
  type GridLayout,
} from "./layout";
import { compileStylesheet, rotationCss } from "./css";
import type { PreviewCell, PreviewSheet, PreviewWorkbook } from "../types";

const BASE_CSS = `
.xpv-root{position:relative;display:flex;flex-direction:column;width:100%;height:100%;min-height:200px;font-family:Calibri,"Segoe UI",system-ui,sans-serif;font-size:11pt;color:#000;background:#fff;overflow:hidden;}
.xpv-tabs{display:flex;gap:2px;align-items:flex-end;padding:4px 6px 0;border-bottom:1px solid #d0d7de;background:#f5f6f7;flex:none;overflow-x:auto;}
.xpv-tab{border:1px solid #d0d7de;border-bottom:none;background:#e9ebee;color:#444;padding:3px 14px;font-size:12px;cursor:pointer;border-radius:3px 3px 0 0;white-space:nowrap;font-family:inherit;}
.xpv-tab.xpv-active{background:#fff;color:#0b57d0;font-weight:600;}
.xpv-viewport{position:relative;flex:1;overflow:auto;background:#fff;}
.xpv-content{position:absolute;top:0;left:0;}
.xpv-layer{position:absolute;}
.xpv-cell{position:absolute;box-sizing:border-box;padding:0 3px;line-height:1.15;white-space:nowrap;overflow:hidden;display:flex;align-items:center;}
.xpv-num{justify-content:flex-end;text-align:right;}
.xpv-text{justify-content:flex-start;text-align:left;}
.xpv-mid{justify-content:center;text-align:center;}
.xpv-spill{overflow:visible;}
.xpv-hcell{position:absolute;box-sizing:border-box;display:flex;align-items:center;justify-content:center;background:#f0f0f0;border-right:1px solid #d0d7de;border-bottom:1px solid #d0d7de;color:#666;font-size:12px;user-select:none;font-family:system-ui,sans-serif;}
.xpv-frozen{background:#fff;}
`;

/**
 * 网格线：以 .xpv-cell 的右/下 border 实现。规则必须注入在 xf 类**之前**——
 * 两者同为同特异性选择器（0,2,0，见下），同源冲突按源序后者胜：前置后带数据
 * 边框的格子由 .xpv-xf-N 覆盖网格线，无 xf 边框的格子吃网格线（见 setSheet
 * 的拼接顺序，曾经后置导致数据边框整体被网格灰覆盖）。
 *
 * 两条规则都按实例作用域前缀（`.xpv-sN …`，见 compileStylesheet）：类名与
 * 网格线开关是每实例一份的，无前缀时同页第二个预览的 <style> 会按源序
 * 覆盖第一个（样式串染），且"关网格线"的实例仍会被另一个实例开着的规则
 * 画上网格线——两种情形都只在多实例共存时出现。
 */
const GRID_COLOR = "#e1e1e1";

/** 实例序号：作用域类 `.xpv-sN` 的唯一来源（模块级自增，进程内不重复）。 */
let instanceSeq = 0;

/** 视口尺寸不可用（happy-dom / 首帧未布局）时的回退渲染尺寸。 */
const FALLBACK_VIEW_W = 800;
const FALLBACK_VIEW_H = 600;

export interface RenderOptions {
  /** 初始 sheet（名称或索引；渲染层内部都走 setSheet 的名称/索引双通道）。 */
  sheet?: string | number;
  showHeaders?: boolean;
  showGridLines?: boolean;
  showTabs?: boolean;
  onSheetChange?: (index: number) => void;
  onRendered?: (durationMs: number) => void;
}

export class SheetRenderer {
  private readonly root: HTMLElement;
  private tabsBar: HTMLElement | null = null;
  private viewport!: HTMLDivElement;
  private content!: HTMLDivElement;
  private cellsLayer!: HTMLDivElement;
  private frozenColLayer!: HTMLDivElement;
  private frozenRowLayer!: HTMLDivElement;
  private frozenCorner!: HTMLDivElement;
  private headerRow!: HTMLDivElement;
  private headerCol!: HTMLDivElement;
  private headerCorner!: HTMLDivElement;
  /** 冻结列的列标 / 冻结行的行号：静态层（不平移），见 updateHeaders。 */
  private colHeadFrozen!: HTMLDivElement;
  private rowHeadFrozen!: HTMLDivElement;
  private styleEl: HTMLStyleElement | null = null;

  private workbook: PreviewWorkbook | null = null;
  private sheetIndex = -1;
  private layout: GridLayout | null = null;
  /** row(0-based) → (col → cell)（每 sheet 重建）。 */
  private rowIndex = new Map<number, Map<number, PreviewCell>>();
  /**
   * 象限前缀 + "r:c" → 元素。四个象限各用互不为前缀的非空前缀（m:/fr:/
   * fc:/fcol:）：updateLayer 的清理以 key.startsWith(prefix) 判归属，主象限
   * 若用空前缀会恒匹配所有键，把冻结象限的元素一并误删、随后全量重建
   * （每滚动帧一次的 DOM churn，曾经的真实缺陷）。
   */
  private rendered = new Map<string, HTMLElement>();
  private renderedHeaders = new Map<string, HTMLElement>();
  private ro: ResizeObserver | null = null;
  private rafId = 0;
  private showHeaders = true;
  private showGridLines = true;
  private showTabs = true;
  private gridOverride: boolean | undefined;
  private opts: RenderOptions = {};
  private destroyFns: (() => void)[] = [];
  private sheetStyleEl: HTMLStyleElement | null = null;
  /** 实例作用域类（`.xpv-sN`）：sheet 级规则的隔离前缀，见 GRID_COLOR 注释。 */
  private readonly scope: string;

  constructor(container: HTMLElement) {
    this.scope = `xpv-s${++instanceSeq}`;
    this.root = document.createElement("div");
    this.root.className = `xpv-root ${this.scope}`;
    container.appendChild(this.root);
  }

  /** 渲染工作簿（首次或换文件）。 */
  render(workbook: PreviewWorkbook, opts: RenderOptions = {}): void {
    this.workbook = workbook;
    this.opts = opts;
    this.showHeaders = opts.showHeaders ?? true;
    this.showTabs = opts.showTabs ?? true;
    this.gridOverride = opts.showGridLines;

    if (!this.styleEl) {
      this.styleEl = document.createElement("style");
      this.styleEl.textContent = BASE_CSS;
      document.head.appendChild(this.styleEl);
      this.destroyFns.push(() => this.styleEl?.remove());
    }
    if (!this.sheetStyleEl) {
      // sheet 级样式（xf 类 + 网格线）独立元素：切 sheet 时整体替换
      this.sheetStyleEl = document.createElement("style");
      document.head.appendChild(this.sheetStyleEl);
      this.destroyFns.push(() => this.sheetStyleEl?.remove());
    }
    this.buildDom();
    // 初始 sheet 越界/缺名时回退 activeSheetIndex（与名称路径的兜底语义一致）。
    // 不回退则 setSheet 提前 return、layout 保持 null，预览静默空白且
    // onParsed/onError 均不触发（此前传越界数字索引的实测行为）。
    if (
      this.setSheet(opts.sheet ?? workbook.activeSheetIndex) < 0 &&
      this.workbook.sheets.length > 0
    ) {
      this.setSheet(workbook.activeSheetIndex);
    }
  }

  private buildDom(): void {
    this.root.textContent = "";

    if (this.showTabs) {
      this.tabsBar = document.createElement("div");
      this.tabsBar.className = "xpv-tabs";
      this.root.appendChild(this.tabsBar);
    }

    this.viewport = document.createElement("div");
    this.viewport.className = "xpv-viewport";
    this.root.appendChild(this.viewport);

    this.content = document.createElement("div");
    this.content.className = "xpv-content";
    this.viewport.appendChild(this.content);

    this.cellsLayer = document.createElement("div");
    this.cellsLayer.className = "xpv-layer";
    this.content.appendChild(this.cellsLayer);

    this.frozenColLayer = document.createElement("div");
    this.frozenRowLayer = document.createElement("div");
    this.frozenCorner = document.createElement("div");
    this.headerRow = document.createElement("div");
    this.headerCol = document.createElement("div");
    this.headerCorner = document.createElement("div");
    this.colHeadFrozen = document.createElement("div");
    this.rowHeadFrozen = document.createElement("div");
    for (const el of [
      this.frozenColLayer,
      this.frozenRowLayer,
      this.frozenCorner,
      this.headerRow,
      this.headerCol,
      this.headerCorner,
      this.colHeadFrozen,
      this.rowHeadFrozen,
    ]) {
      el.classList.add("xpv-layer");
      this.viewport.appendChild(el);
    }
    this.frozenColLayer.classList.add("xpv-frozen");
    this.frozenRowLayer.classList.add("xpv-frozen");
    this.frozenCorner.classList.add("xpv-frozen");

    const onScroll = () => {
      if (this.rafId) return;
      this.rafId = requestAnimationFrame(() => {
        this.rafId = 0;
        this.updateOverlays();
      });
    };
    this.viewport.addEventListener("scroll", onScroll, { passive: true });
    this.destroyFns.push(() =>
      this.viewport.removeEventListener("scroll", onScroll),
    );

    if (typeof ResizeObserver !== "undefined") {
      this.ro = new ResizeObserver(() => this.updateOverlays(true));
      this.ro.observe(this.viewport);
      this.destroyFns.push(() => this.ro?.disconnect());
    }
  }

  /** 切换 sheet（名称或索引）。返回实际索引；无效返回 -1。 */
  setSheet(nameOrIndex: string | number): number {
    if (!this.workbook) return -1;
    const idx =
      typeof nameOrIndex === "number"
        ? nameOrIndex
        : this.workbook.sheets.findIndex((s) => s.name === nameOrIndex);
    if (idx < 0 || idx >= this.workbook.sheets.length) return -1;
    const sheet = this.workbook.sheets[idx];
    this.sheetIndex = idx;
    this.showGridLines = this.gridOverride ?? sheet.showGridLines;

    const t0 = performance.now();
    this.layout = buildLayout(sheet);
    this.rowIndex.clear();
    for (const r of sheet.rows) {
      const m = new Map<number, PreviewCell>();
      for (const c of r.cells) m.set(c.col, c);
      this.rowIndex.set(r.index - 1, m);
    }
    // 网格线规则前置于 xf 类（级联依据见 GRID_COLOR 注释），数据边框才能
    // 在同特异性下按源序覆盖网格线；两条规则都带实例作用域前缀，多实例
    // 共存时才不会互相串样式。
    this.sheetStyleEl!.textContent =
      (this.showGridLines
        ? `.${this.scope} .xpv-cell{border-right:1px solid ${GRID_COLOR};border-bottom:1px solid ${GRID_COLOR};}`
        : "") + compileStylesheet(sheet.styles, this.scope);

    this.rendered.clear();
    this.renderedHeaders.clear();
    for (const el of [
      this.cellsLayer,
      this.frozenColLayer,
      this.frozenRowLayer,
      this.frozenCorner,
      this.headerRow,
      this.headerCol,
      this.headerCorner,
      this.colHeadFrozen,
      this.rowHeadFrozen,
    ]) {
      el.textContent = "";
    }
    this.viewport.scrollTop = 0;
    this.viewport.scrollLeft = 0;

    this.rebuildGeometry();
    this.renderTabs();
    this.updateOverlays(true);
    this.opts.onRendered?.(performance.now() - t0);
    return idx;
  }

  private renderTabs(): void {
    if (!this.tabsBar || !this.workbook) return;
    this.tabsBar.textContent = "";
    this.workbook.sheets.forEach((s, i) => {
      if (!s.visible) return;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "xpv-tab" + (i === this.sheetIndex ? " xpv-active" : "");
      btn.textContent = s.name || `Sheet${i + 1}`;
      btn.addEventListener("click", () => {
        if (i !== this.sheetIndex) {
          this.setSheet(i);
          this.opts.onSheetChange?.(i);
        }
      });
      this.tabsBar!.appendChild(btn);
    });
  }

  private rebuildGeometry(): void {
    const l = this.layout!;
    const headW = this.showHeaders ? HEADER_COL_PX : 0;
    const headH = this.showHeaders ? HEADER_ROW_PX : 0;
    // 滚动尺寸 = 表头 + 全网格
    this.content.style.width = `${headW + l.totalWidth}px`;
    this.content.style.height = `${headH + l.totalHeight}px`;

    // 各层（网格坐标，格子直接放网格坐标；层本身相对 viewport 定位）
    this.cellsLayer.style.left = `${headW}px`;
    this.cellsLayer.style.top = `${headH}px`;
    this.cellsLayer.style.width = `${l.totalWidth}px`;
    this.cellsLayer.style.height = `${l.totalHeight}px`;

    this.frozenColLayer.style.left = `${headW}px`;
    this.frozenColLayer.style.top = `${headH}px`;
    this.frozenColLayer.style.width = `${l.colLeft[l.frozenCols]}px`;
    this.frozenColLayer.style.height = `${l.totalHeight}px`;
    this.frozenColLayer.style.overflow = "hidden";
    this.frozenColLayer.style.zIndex = "3";

    this.frozenRowLayer.style.left = `${headW}px`;
    this.frozenRowLayer.style.top = `${headH}px`;
    this.frozenRowLayer.style.width = `${l.totalWidth}px`;
    this.frozenRowLayer.style.height = `${l.rowTop[l.frozenRows]}px`;
    this.frozenRowLayer.style.overflow = "hidden";
    this.frozenRowLayer.style.zIndex = "4";

    this.frozenCorner.style.left = `${headW}px`;
    this.frozenCorner.style.top = `${headH}px`;
    this.frozenCorner.style.width = `${l.colLeft[l.frozenCols]}px`;
    this.frozenCorner.style.height = `${l.rowTop[l.frozenRows]}px`;
    this.frozenCorner.style.overflow = "hidden";
    this.frozenCorner.style.zIndex = "6";

    this.headerRow.style.left = `${headW}px`;
    this.headerRow.style.top = "0px";
    this.headerRow.style.width = `${l.totalWidth}px`;
    this.headerRow.style.height = `${headH}px`;
    this.headerRow.style.overflow = "hidden";
    this.headerRow.style.zIndex = "5";

    this.headerCol.style.left = "0px";
    this.headerCol.style.top = `${headH}px`;
    this.headerCol.style.width = `${headW}px`;
    this.headerCol.style.height = `${l.totalHeight}px`;
    this.headerCol.style.overflow = "hidden";
    this.headerCol.style.zIndex = "5";

    this.headerCorner.style.left = "0px";
    this.headerCorner.style.top = "0px";
    this.headerCorner.style.width = `${headW}px`;
    this.headerCorner.style.height = `${headH}px`;
    this.headerCorner.style.zIndex = "7";

    // 冻结列的列标带 / 冻结行的行号带：几何在网格坐标系，由 updateOverlays
    // 双向钉住（translate(scrollLeft, scrollTop)），与 headerRow/headerCol 的
    // 单轴钉住带互补。背景同 .xpv-hcell 的 #f0f0f0——层必须不透明，否则下方
    // 滚动带里滑入冻结带的列标/行号会从隐藏列的缝隙透出。
    this.colHeadFrozen.style.left = `${headW}px`;
    this.colHeadFrozen.style.top = "0px";
    this.colHeadFrozen.style.width = `${l.colLeft[l.frozenCols]}px`;
    this.colHeadFrozen.style.height = `${headH}px`;
    this.colHeadFrozen.style.overflow = "hidden";
    this.colHeadFrozen.style.zIndex = "6";
    this.colHeadFrozen.style.background = "#f0f0f0";
    this.rowHeadFrozen.style.left = "0px";
    this.rowHeadFrozen.style.top = `${headH}px`;
    this.rowHeadFrozen.style.width = `${headW}px`;
    this.rowHeadFrozen.style.height = `${l.rowTop[l.frozenRows]}px`;
    this.rowHeadFrozen.style.overflow = "hidden";
    this.rowHeadFrozen.style.zIndex = "6";
    this.rowHeadFrozen.style.background = "#f0f0f0";

    if (this.showHeaders) {
      this.headerCorner.classList.add("xpv-hcell");
      this.headerCorner.style.display = "";
      this.colHeadFrozen.style.display = "";
      this.rowHeadFrozen.style.display = "";
    } else {
      this.headerCorner.classList.remove("xpv-hcell");
      for (const el of [
        this.headerRow,
        this.headerCol,
        this.headerCorner,
        this.colHeadFrozen,
        this.rowHeadFrozen,
      ]) {
        el.style.display = "none";
      }
    }
  }

  /** 滚动/尺寸变化：更新覆盖层 transform 并做窗口差量。 */
  private updateOverlays(force = false): void {
    const l = this.layout;
    if (!l) return;
    const st = this.viewport.scrollTop;
    const sl = this.viewport.scrollLeft;
    if (!force && st === this.lastSt && sl === this.lastSl) return;
    this.lastSt = st;
    this.lastSl = sl;

    // transform 只抵消"需要钉住"的轴：钉住的轴用 +scroll，需要跟随滚动的
    // 轴不加 transform。覆盖层本身随滚动位移（见文件头不变量），此前按"绝对
    // 定位子元素不随滚动"写成了 -scroll/错轴，实测滚动时表头与冻结条带以
    // 两倍速度滑走（真实浏览器几何实测）。
    const pin = `translate(${sl}px, ${st}px)`;
    this.frozenColLayer.style.transform = `translateX(${sl}px)`;
    this.frozenRowLayer.style.transform = `translateY(${st}px)`;
    this.headerRow.style.transform = `translateY(${st}px)`;
    this.headerCol.style.transform = `translateX(${sl}px)`;
    this.frozenCorner.style.transform = pin;
    this.headerCorner.style.transform = pin;
    this.colHeadFrozen.style.transform = pin;
    this.rowHeadFrozen.style.transform = pin;

    this.updateWindow();
  }
  private lastSt = 0;
  private lastSl = 0;

  private updateWindow(): void {
    const l = this.layout;
    if (!l || !this.workbook) return;
    const headW = this.showHeaders ? HEADER_COL_PX : 0;
    const headH = this.showHeaders ? HEADER_ROW_PX : 0;
    // 主区窗口：可视区在"网格坐标"下的位置 = 滚动偏移 − 表头（表头是 overlay，
    // 占据视口顶部/左侧但不在滚动坐标里体现）+ 被冻结层覆盖的尺寸；可用宽高
    // 相应扣除表头与冻结区。此前漏减表头偏移，窗口整体偏 24/40px，靠 200px
    // buffer 掩盖（表头加高或 buffer 调小时会露出顶部缺行）。
    const vw =
      (this.viewport.clientWidth || FALLBACK_VIEW_W) -
      headW -
      l.colLeft[l.frozenCols];
    const vh =
      (this.viewport.clientHeight || FALLBACK_VIEW_H) -
      headH -
      l.rowTop[l.frozenRows];
    if (vw <= 0 || vh <= 0) return;

    const main = computeVisibleRange(
      l,
      this.viewport.scrollTop - headH + l.rowTop[l.frozenRows],
      this.viewport.scrollLeft - headW + l.colLeft[l.frozenCols],
      vw,
      vh,
    );

    const mainRange = {
      rowStart: Math.max(main.rowStart, l.frozenRows),
      rowEnd: main.rowEnd,
      colStart: Math.max(main.colStart, l.frozenCols),
      colEnd: main.colEnd,
    };
    // 与主区窗口相交的合并主格：窗口外的主格也要渲染（合并显示完整性），
    // 但只召回主格本身、不扩大窗口与表头范围（整列合并的 DOM 爆炸，见
    // visibleMergeAnchors 注释）
    const anchors = visibleMergeAnchors(l, mainRange);

    this.updateLayer(this.cellsLayer, mainRange, "m:", anchors);

    if (l.frozenRows > 0) {
      this.updateLayer(
        this.frozenRowLayer,
        {
          rowStart: 0,
          rowEnd: l.frozenRows,
          colStart: Math.max(main.colStart, l.frozenCols),
          colEnd: main.colEnd,
        },
        "fr:",
      );
      this.updateLayer(
        this.frozenCorner,
        {
          rowStart: 0,
          rowEnd: l.frozenRows,
          colStart: 0,
          colEnd: l.frozenCols,
        },
        "fc:",
      );
    }
    if (l.frozenCols > 0) {
      this.updateLayer(
        this.frozenColLayer,
        {
          rowStart: Math.max(main.rowStart, l.frozenRows),
          rowEnd: main.rowEnd,
          colStart: 0,
          colEnd: l.frozenCols,
        },
        "fcol:",
      );
    }

    this.updateHeaders(main);
  }

  /** 单象限差量渲染。prefix 区分象限 key 命名空间；anchors 为需增补召回的
   * 合并主格（在窗口外，见 visibleMergeAnchors），仅主象限使用。 */
  private updateLayer(
    layer: HTMLElement,
    range: {
      rowStart: number;
      rowEnd: number;
      colStart: number;
      colEnd: number;
    },
    prefix: string,
    anchors: { row: number; col: number }[] = [],
  ): void {
    const l = this.layout!;
    const want = new Set<string>();

    // 单格渲染（含合并几何与隐藏行列检查）：窗口循环与 anchor 召回共用
    const ensure = (r: number, cell: PreviewCell): void => {
      const c = cell.col;
      const merge = l.mergeByAnchor.get(`${r}:${c}`);
      const left = l.colLeft[c];
      const top = l.rowTop[r];
      const width = merge
        ? l.colLeft[Math.min(c + merge.colSpan, l.colCount)] - left
        : l.colWidths[c];
      const height = merge
        ? l.rowTop[Math.min(r + merge.rowSpan, l.rowCount)] - top
        : l.rowHeights[r];
      if (width <= 0 || height <= 0) return; // 隐藏行列中的格子不渲染
      // 文本溢出：溢出宽度封顶到行内下一个"有内容/被合并覆盖"格的前缘
      // （Excel 溢出到第一个非空格前截断——此前无上限，长文本会视觉盖到
      // 远处内容格上）。formulaStr 是公式的缓存字符串结果，与文本同属
      // 左对齐类（见 cellClasses），Excel 同样溢出到相邻空格。
      let spillPx: number | undefined;
      if (cell.type === "string" || cell.type === "formulaStr") {
        const limit = this.spillLimitPx(c, r, range.colEnd);
        if (limit > width) spillPx = limit;
      }
      const key = `${prefix}${r}:${c}`;
      want.add(key);
      if (!this.rendered.has(key)) {
        this.createCell(
          layer,
          key,
          cell,
          this.workbook!.sheets[this.sheetIndex],
          left,
          top,
          width,
          height,
          spillPx,
        );
      }
    };

    for (let r = range.rowStart; r < range.rowEnd; r++) {
      const rowMap = this.rowIndex.get(r);
      if (!rowMap) continue;
      for (const cell of rowMap.values()) {
        if (cell.col < range.colStart || cell.col >= range.colEnd) continue;
        if (l.coveredBy.has(`${r}:${cell.col}`)) continue;
        ensure(r, cell);
      }
    }
    for (const a of anchors) {
      // anchor 正常不是被覆盖格（coveredBy 不含 anchor 自身）；重叠合并的
      // 异常文件里后写覆盖先写，防御性跳过
      if (l.coveredBy.has(`${a.row}:${a.col}`)) continue;
      const cell = this.rowIndex.get(a.row)?.get(a.col);
      if (cell) ensure(a.row, cell);
    }

    for (const [key, el] of this.rendered) {
      if (!key.startsWith(prefix)) continue;
      if (!want.has(key)) {
        el.remove();
        this.rendered.delete(key);
      }
    }
  }

  private createCell(
    layer: HTMLElement,
    key: string,
    cell: PreviewCell,
    sheet: PreviewSheet,
    left: number,
    top: number,
    width: number,
    height: number,
    spillPx?: number,
  ): void {
    const el = document.createElement("div");
    el.className = this.cellClasses(cell, sheet);
    el.style.left = `${left}px`;
    el.style.top = `${top}px`;
    el.style.width = `${width}px`;
    el.style.height = `${height}px`;

    const wb = this.workbook!;
    const xf =
      cell.styleIndex != null ? sheet.styles.xfs[cell.styleIndex] : undefined;
    const fmt = formatCellValue(
      cell.type,
      cell.value,
      xf?.numFmtCode ?? "General",
      wb.dateSystem,
    );
    if (fmt.color) el.style.color = fmt.color;
    const rotation = rotationCss(xf?.alignment?.textRotation);
    if (fmt.text !== "") {
      if (spillPx !== undefined && !rotation) {
        // 溢出截断容器：el 本体保持原格宽（背景/边框不外溢、overflow 放开），
        // 文本在 span 里按溢出上限截断（Excel 溢出到第一个非空格前）
        el.classList.add("xpv-spill");
        const span = document.createElement("span");
        span.textContent = fmt.text;
        span.style.cssText = `max-width:${spillPx}px;overflow:hidden;`;
        el.appendChild(span);
      } else if (rotation) {
        const span = document.createElement("span");
        span.textContent = fmt.text;
        span.style.cssText = rotation;
        el.appendChild(span);
      } else {
        el.textContent = fmt.text;
      }
    }
    layer.appendChild(el);
    this.rendered.set(key, el);
  }

  private cellClasses(cell: PreviewCell, sheet: PreviewSheet): string {
    const xf =
      cell.styleIndex != null ? sheet.styles.xfs[cell.styleIndex] : undefined;
    let cls = "xpv-cell";
    // General 对齐分流：数字右、文本左、布尔/错误/公式串居中（Excel 语义）
    if (!xf?.alignment?.horizontal) {
      if (cell.type === "number") cls += " xpv-num";
      else if (cell.type === "string" || cell.type === "formulaStr")
        cls += " xpv-text";
      else cls += " xpv-mid";
    }
    if (xf) cls += ` xpv-xf-${cell.styleIndex}`;
    return cls;
  }

  /**
   * 文本溢出上限（px）：本格左缘到行内下一个"有内容 / 被合并覆盖"格左缘的
   * 距离；无阻挡时到扫描终点。扫描以窗口右界 +1 为上限——溢出超出视口的
   * 部分本就会被 viewport 裁剪，无需扫到网格尽头（避免超宽行 O(总列数)
   * 的逐格扫描）。
   */
  private spillLimitPx(col: number, row: number, colEnd: number): number {
    const l = this.layout!;
    const rowMap = this.rowIndex.get(row);
    const stop = Math.min(l.colCount, Math.max(colEnd + 1, col + 2));
    for (let c = col + 1; c < stop; c++) {
      if (rowMap?.has(c) || l.coveredBy.has(`${row}:${c}`)) {
        return l.colLeft[c] - l.colLeft[col];
      }
    }
    return l.colLeft[stop] - l.colLeft[col];
  }

  private updateHeaders(range: {
    rowStart: number;
    rowEnd: number;
    colStart: number;
    colEnd: number;
  }): void {
    if (!this.showHeaders) return;
    const l = this.layout!;
    const want = new Set<string>();

    // 列标拆两条带：c < frozenCols 进静态层（冻结列恒可见，Excel 中冻结列
    // 的列标固定不动——此前混入 headerRow 平移带，横向滚动时会滑入表头角
    // 块下方消失）；c >= frozenCols 进平移带随内容滚动。
    for (let c = 0; c < l.frozenCols; c++) {
      this.headerCell(
        want,
        `fh:${c}`,
        this.colHeadFrozen,
        columnIndexLabel(c),
        {
          left: `${l.colLeft[c]}px`,
          width: `${l.colWidths[c]}px`,
          height: `${HEADER_ROW_PX}px`,
          top: "0px",
        },
      );
    }
    const colStart = Math.max(range.colStart, l.frozenCols);
    for (let c = colStart; c < range.colEnd; c++) {
      this.headerCell(want, `h:${c}`, this.headerRow, columnIndexLabel(c), {
        left: `${l.colLeft[c]}px`,
        width: `${l.colWidths[c]}px`,
        height: `${HEADER_ROW_PX}px`,
        top: "0px",
      });
    }

    // 行号同理拆静态/平移两条带
    for (let r = 0; r < l.frozenRows; r++) {
      this.headerCell(want, `fv:${r}`, this.rowHeadFrozen, String(r + 1), {
        top: `${l.rowTop[r]}px`,
        height: `${l.rowHeights[r]}px`,
        width: `${HEADER_COL_PX}px`,
        left: "0px",
      });
    }
    const rowStart = Math.max(range.rowStart, l.frozenRows);
    for (let r = rowStart; r < range.rowEnd; r++) {
      this.headerCell(want, `v:${r}`, this.headerCol, String(r + 1), {
        top: `${l.rowTop[r]}px`,
        height: `${l.rowHeights[r]}px`,
        width: `${HEADER_COL_PX}px`,
        left: "0px",
      });
    }

    for (const [key, el] of this.renderedHeaders) {
      if (!want.has(key)) {
        el.remove();
        this.renderedHeaders.delete(key);
      }
    }
  }

  /** 表头格子的创建/复用（四条表头带共用）。几何与内容对同一 key 恒定
   * （列标由列号、行号由行号唯一决定），按文件头不变量只在创建时写入，
   * 复用路径零 DOM 写入。 */
  private headerCell(
    want: Set<string>,
    key: string,
    layer: HTMLElement,
    label: string,
    pos: { left: string; top: string; width: string; height: string },
  ): void {
    want.add(key);
    let el = this.renderedHeaders.get(key);
    if (!el) {
      el = document.createElement("div");
      el.className = "xpv-hcell";
      el.textContent = label;
      el.style.left = pos.left;
      el.style.top = pos.top;
      el.style.width = pos.width;
      el.style.height = pos.height;
      layer.appendChild(el);
      this.renderedHeaders.set(key, el);
    }
  }

  destroy(): void {
    if (this.rafId) cancelAnimationFrame(this.rafId);
    for (const fn of this.destroyFns) fn();
    this.destroyFns = [];
    this.rendered.clear();
    this.renderedHeaders.clear();
    this.ro = null;
    this.root.remove();
  }

  get activeSheetIndex(): number {
    return this.sheetIndex;
  }
}
