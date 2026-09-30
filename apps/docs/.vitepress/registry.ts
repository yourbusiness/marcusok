import excelExporterPkg from "@marcusok/excel-exporter/package.json" with { type: "json" };
import excelPreviewPkg from "@marcusok/excel-preview/package.json" with { type: "json" };
import progressOverlayPkg from "@marcusok/progress-overlay/package.json" with { type: "json" };

export interface LocalizedText {
  zh: string;
  en: string;
}

/** One number card on the home page stats block, contributed by a package. */
export interface HomeStat {
  key: string;
  value: number;
  decimals: number;
  zh: string;
  en: string;
  suffix?: string;
  /** Optional external link; rendered as a clickable stat card. */
  href?: string;
}

/** One feature card on the home page highlights section. */
export interface PackageHighlight {
  /** Icon key from the home highlights icon set (see PackageHighlights.vue);
   *  unrecognized values render as the raw string (emoji fallback). */
  icon: string;
  title: LocalizedText;
  details: LocalizedText;
}

/** A sidebar group for a package; `id` is the sub-directory under packages/<dir>/. */
export interface PackageSection {
  id: string;
  label: LocalizedText;
  collapsed?: boolean;
}

/** A dist asset (wasm, worker, ...) copied into public/ at build time. */
export interface RuntimeAsset {
  resolveFrom: string;
  /**
   * Optional: resolve `resolveFrom` within another package's dependency
   * context (via createRequire). Use this when the asset ships with a
   * transitive dependency rather than a direct one, so the docs app does not
   * need to list the asset's source package as a direct dependency itself.
   */
  through?: string;
  /** File path within the resolved package directory. */
  file: string;
  /** Destination path relative to the docs public/ directory. */
  to: string;
}

/** One series (legend entry + bar color) in a benchmark chart. */
export interface BenchmarkSeriesDef {
  /** Unique key within this chart; must match a key in BenchmarkBar.values. */
  key: string;
  /** Localized legend label. */
  label: LocalizedText;
  /** CSS color for bars; falls back to --vp-c-brand-1 for the first series. */
  color?: string;
}

/** One bar group on the x-axis. */
export interface BenchmarkBar {
  /** X-axis label, e.g. "10k". */
  label: string;
  /** Map of seriesKey -> value (typically milliseconds). */
  values: Record<string, number>;
}

/**
 * A benchmark chart dataset contributed by a package; rendered on the home
 * page and on the package's performance page. The data model is generic:
 * any number of series can be compared (main vs stream, v1 vs v2, etc.).
 */
export interface BenchmarkSeries {
  source: LocalizedText;
  /** Bar groups along the x-axis. */
  data: BenchmarkBar[];
  /** Series definitions (legend + color). Must have at least 1 entry. */
  series: BenchmarkSeriesDef[];
}

/**
 * 包大类：导出（现有 excel-exporter，后续或补其他文档类型导出）/
 * 文档预览 / 公共组件（业务包共用的底层能力，如进度遮罩）。分类只是注册表
 * 里的元数据，不落物理目录——目录结构仍按 packages/<name> 平铺，避免破坏
 * 各处按目录名工作的链接与约定。
 */
export type PackageCategory = "export" | "preview" | "shared";

export interface PackageCategoryDef {
  id: PackageCategory;
  label: LocalizedText;
}

/** 全部分类（含展示顺序与双语文案）：导航、侧边栏与首页卡片共用。 */
export const PACKAGE_CATEGORIES: PackageCategoryDef[] = [
  { id: "export", label: { zh: "导出", en: "Export" } },
  { id: "preview", label: { zh: "文档预览", en: "Document Preview" } },
  { id: "shared", label: { zh: "公共组件", en: "Shared Components" } },
];

/**
 * 按大类聚包（分类保持 PACKAGE_CATEGORIES 顺序，包保持注册顺序）。
 * 没有包的大类不产出分组——尚无交付包的分类不渲染空标题。
 */
export function groupPackagesByCategory(
  pkgs: readonly PackageEntry[],
): { category: PackageCategoryDef; pkgs: PackageEntry[] }[] {
  return PACKAGE_CATEGORIES.flatMap((category) => {
    const matched = pkgs.filter((p) => p.category === category.id);
    return matched.length > 0 ? [{ category, pkgs: matched }] : [];
  });
}

export interface PackageEntry {
  /** Directory name under the docs root: packages/<dir>/*.md */
  dir: string;
  npmName: string;
  /** 包所属大类：导航、侧边栏与首页卡片按它分组（见 PACKAGE_CATEGORIES）。 */
  category: PackageCategory;
  version: string;
  status: "stable" | "beta" | "alpha";
  tagline: LocalizedText;
  keywords: string[];
  /**
   * Whether this package ships a zh/ doc mirror. Single source of truth
   * for nav/sidebar/home visibility in the Chinese locale. Validated
   * against the actual zh/packages/<dir>/ directory by validateDocsTree
   * (config.ts) at build time, so the flag and dir can never drift.
   */
  zh: boolean;
  /**
   * Optional extra sidebar groups besides the default guide/examples/api.
   * `id` must be a directory under packages/<dir>/ containing the markdown.
   */
  sections?: PackageSection[];
  /** Optional home-page stat cards contributed by this package (keys must be globally unique). */
  homeStats?: HomeStat[];
  /** Optional home-page highlight cards contributed by this package. */
  highlights?: PackageHighlight[];
  /** Dist assets (wasm, worker) copied into public/ at build time. */
  runtimeAssets?: RuntimeAsset[];
  /** Home-page benchmark chart datasets (one SVG block per series). */
  benchmarks?: BenchmarkSeries[];
  /**
   * Name of a globally-registered Vue component (from theme/components/)
   * to render as this package's live demo. Referenced via <PackageDemo>.
   */
  demo?: string;
}

/** Default sidebar groups used when a package does not declare `sections`. */
export const DEFAULT_PACKAGE_SECTIONS: PackageSection[] = [
  { id: "guide", label: { zh: "指南", en: "Guide" }, collapsed: false },
  {
    id: "examples",
    label: { zh: "使用案例", en: "Examples" },
    collapsed: true,
  },
  {
    id: "api",
    label: { zh: "API 参考", en: "API Reference" },
    collapsed: true,
  },
];

/**
 * Resolve the effective sidebar sections for a package: the default
 * guide/examples/api groups first, then the package's own `sections`.
 * A custom section with the same `id` as a default overrides it (e.g. to
 * relabel "Guide"), so packages only need to declare *extra* groups.
 */
export function resolvePackageSections(p: PackageEntry): PackageSection[] {
  const merged = new Map<string, PackageSection>();
  for (const s of [...DEFAULT_PACKAGE_SECTIONS, ...(p.sections ?? [])]) {
    merged.set(s.id, s);
  }
  return [...merged.values()];
}

/**
 * Home-page stats: package count (always first) followed by the given
 * packages' declared `homeStats`. Defaults to the full registry; callers on
 * the Chinese site pass the locale-filtered list so en-only packages do not
 * leak stats onto the zh home page. Keys are guaranteed unique.
 */
export function getAllHomeStats(
  pkgs: readonly PackageEntry[] = packages,
): HomeStat[] {
  const primaryPackage = pkgs[0];
  const npmScope = primaryPackage?.npmName.split("/")[0]?.replace(/^@/, "");
  return [
    {
      key: "packages",
      value: pkgs.length,
      decimals: 0,
      // 口径是"文档站收录的包"而非"npm 已发布"：@marcusok/xlsx-core 已发
      // npm 但有意不在本注册表（共享引擎层，其用法随两个业务包的文档讲），
      // 按发布数会虚报（对照 guide/index.md 的仓库结构表则含它）
      zh: "文档站收录库包",
      en: "Documented packages",
      href:
        pkgs.length === 1
          ? `https://www.npmjs.com/package/${primaryPackage?.npmName ?? ""}`
          : npmScope
            ? `https://www.npmjs.com/org/${npmScope}`
            : undefined,
    },
    ...pkgs.flatMap((p) => p.homeStats ?? []),
  ];
}

/**
 * Package registry — the single source of truth for the docs site.
 * Adding a new package: add it to apps/docs/package.json dependencies,
 * create packages/<dir>/ markdown, then append one entry here (declaring
 * its `category`). Sidebar, nav, home cards, highlights and stats are
 * generated from this list. Version is read from the package's own
 * package.json (single source of truth).
 */
export const packages: PackageEntry[] = [
  {
    dir: "excel-exporter",
    npmName: "@marcusok/excel-exporter",
    category: "export",
    version: excelExporterPkg.version,
    status: "stable",
    zh: true,
    tagline: {
      zh: "Excel 导出引擎（WASM + Fast stream，10 万行约 0.8s）",
      en: "Excel export engine (WASM + Fast stream, ~0.8s at 100k rows)",
    },
    keywords: ["excel", "xlsx", "export", "wasm"],
    // No runtimeAssets: the wasm/worker ship with the package and are located
    // by default (new URL(<file>, import.meta.url) next to the entry), which
    // the bundler emits as base-aware hashed assets — nothing to copy into
    // public/. The runtimeAssets mechanism stays for future packages that
    // genuinely need public-dir copies.
    benchmarks: [
      {
        data: [
          { label: "10k", values: { auto: 120 } },
          { label: "50k", values: { auto: 400 } },
          { label: "100k", values: { auto: 780 } },
        ],
        series: [
          {
            key: "auto",
            label: { zh: "auto 路径", en: "Auto path" },
          },
        ],
        source: {
          zh: "本机实测：真实 Chrome，6 列混合类型，auto 路径（毫秒）。在线演示的 sales 数据集为 9 列，耗时不能与此口径直接对照。",
          en: "Measured locally: real Chrome, 6 mixed-type columns, auto path (ms). The live demo's sales dataset has 9 columns; its timings are not directly comparable to this baseline.",
        },
      },
    ],
    homeStats: [
      {
        key: "rows",
        value: 0.78,
        decimals: 2,
        zh: "10 万行导出耗时",
        en: "100k rows export time",
        suffix: "s",
      },
      {
        key: "modes",
        value: 4,
        decimals: 0,
        zh: "导出模式",
        en: "Export modes",
      },
      {
        // Keep in sync with StylePresets in packages/excel-exporter/src/style-presets.ts
        // (8 since `bordered` was added; the value is hand-maintained because the
        // registry must stay importable from the browser bundle).
        key: "presets",
        value: 8,
        decimals: 0,
        zh: "内置样式预设",
        en: "Style presets",
      },
    ],
    highlights: [
      {
        icon: "zap",
        title: { zh: "性能", en: "Performance" },
        details: {
          zh: "Fast stream 核心，10 万行导出约 0.8s（本机实测）；Worker 多线程避免长时间占用主线程。",
          en: "Fast stream core exports 100k rows in ~0.8s (measured locally); Web Workers keep heavy work off the main thread.",
        },
      },
      {
        icon: "pen",
        title: { zh: "声明式", en: "Declarative" },
        details: {
          zh: "用配置描述列、样式与格式，一行代码完成导出，不必手写单元格与样式对象。",
          en: "Describe columns, styles and formats with plain config; export with one call.",
        },
      },
      {
        icon: "route",
        title: { zh: "自动路由", en: "Auto Routing" },
        details: {
          zh: "自动模式根据数据量选择最优路径，数据量变化时业务代码零改动。",
          en: "Auto mode picks the best path by row count; business code never changes as data grows.",
        },
      },
      {
        icon: "shield",
        title: { zh: "多级兜底", en: "Layered Fallbacks" },
        details: {
          zh: "环境不支持或 WASM 加载失败时自动降级到无样式纯 JS 快速流，多数异常下仍能拿到导出文件。",
          en: "Automatically degrades to the style-less pure-JS fast stream when WASM is unavailable, so most failures still produce a file.",
        },
      },
    ],
    demo: "ExportDemo",
  },
  {
    dir: "excel-preview",
    npmName: "@marcusok/excel-preview",
    category: "preview",
    version: excelPreviewPkg.version,
    // 已发 1.0.x 正式版（initial release + patch），与导出包同为 stable 口径
    status: "stable",
    zh: true,
    tagline: {
      zh: "xlsx 只读预览（Worker 解析 + 虚拟滚动，样式/合并/冻结/数字格式还原）",
      en: "Read-only xlsx preview (worker parsing + virtual scrolling, styles/merges/freeze/number formats restored)",
    },
    keywords: ["excel", "xlsx", "preview", "viewer", "virtual-scroll"],
    homeStats: [
      {
        key: "preview-rows",
        value: 100,
        decimals: 0,
        zh: "预览行数上限量级",
        en: "Preview scale (rows)",
        suffix: "k",
      },
    ],
    highlights: [
      {
        icon: "zap",
        title: { zh: "大文件不卡", en: "Large files stay smooth" },
        details: {
          zh: "解析在 Web Worker 内完成（实测 10 万行 × 10 列约 1.5s，全程不冻结 UI）；渲染层 DOM 虚拟滚动只渲染视口内格子。",
          en: "Parsing runs in a Web Worker (measured ~1.5s for 100k rows × 10 cols, UI never freezes); the DOM renderer virtualizes to viewport-only cells.",
        },
      },
      {
        icon: "pen",
        title: { zh: "样式还原", en: "Style fidelity" },
        details: {
          zh: "字体/填充/边框/对齐/数字格式/合并/冻结/隐藏行列；主题色（theme+tint）与 indexed 色经自研覆盖层找回（解析引擎读取侧会丢弃它们）。",
          en: "Fonts, fills, borders, alignment, number formats, merges, frozen panes, hidden rows/columns; theme and indexed colors are recovered by a dedicated overlay layer (the engine drops them on read).",
        },
      },
      {
        icon: "shield",
        title: { zh: "数字格式忠实", en: "Number-format fidelity" },
        details: {
          zh: "内置 id 用 Excel 实际行为表（非 ECMA 标准串）；负号/会计括号/货币字面量/累计时长/分钟邻接等引擎缺陷全部在补偿层修复。",
          en: "Built-in format ids use Excel's actual behavior table (not the ECMA strings); engine defects — negative signs, accounting parens, currency literals, elapsed durations, minute adjacency — are compensated in a dedicated layer.",
        },
      },
      {
        icon: "route",
        title: { zh: "框架无关", en: "Framework-agnostic" },
        details: {
          zh: "纯 TS 核心 + createPreview 一行接入；低层 parseWorkbookBytes 出纯数据模型，供 React/Vue 薄封装或 SSR 消费。",
          en: "Pure TypeScript core with a one-line createPreview; the low-level parseWorkbookBytes emits a plain data model for React/Vue wrappers or SSR.",
        },
      },
    ],
    demo: "PreviewDemo",
  },
  {
    dir: "progress-overlay",
    npmName: "@marcusok/progress-overlay",
    category: "shared",
    version: progressOverlayPkg.version,
    status: "stable",
    zh: true,
    tagline: {
      zh: "通用进度遮罩（旋转圆环 / 百分比条，毛玻璃面板，明暗主题）",
      en: "Generic progress overlay (spinner / percentage bar, glass panel, light & dark themes)",
    },
    keywords: ["overlay", "progress", "spinner", "loading", "ui"],
    highlights: [
      {
        icon: "zap",
        title: { zh: "零依赖", en: "Zero dependencies" },
        details: {
          zh: "纯 DOM + 一次注入的样式，无框架绑定；React/Vue/原生页面都能一行接入。",
          en: "Pure DOM with a single injected stylesheet, no framework tie-in; drop into React, Vue or vanilla pages alike.",
        },
      },
      {
        icon: "pen",
        title: { zh: "两态自适应", en: "Two-state adaptive" },
        details: {
          zh: "时长未知显示旋转圆环，流式进度一到即切百分比条；进度粒度由调用方数据源决定。",
          en: "Spinner while duration is unknown, switching to a percentage bar as soon as streamed progress arrives; granularity follows the caller's data source.",
        },
      },
      {
        icon: "shield",
        title: { zh: "并发与无障碍", en: "Concurrency & a11y" },
        details: {
          zh: "并发任务共用一份 DOM、引用计数收尾；aria-live 文案区与 progressbar 语义、reduced-motion 全套支持。",
          en: "Concurrent tasks share one reference-counted DOM node; aria-live labels, progressbar semantics and reduced-motion are all wired in.",
        },
      },
    ],
  },
];
