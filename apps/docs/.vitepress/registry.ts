import excelExporterPkg from "@marcusok/excel-exporter/package.json" with { type: "json" };
import excelPreviewPkg from "@marcusok/excel-preview/package.json" with { type: "json" };
import wordPreviewPkg from "@marcusok/word-preview/package.json" with { type: "json" };
import pdfPreviewPkg from "@marcusok/pdf-preview/package.json" with { type: "json" };
import pptPreviewPkg from "@marcusok/ppt-preview/package.json" with { type: "json" };

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

/** A single-file dist asset (wasm, worker, ...) copied into public/ at build
 *  time. `file` resolves relative to the directory of `resolveFrom`'s entry
 *  (e.g. a wasm sitting next to the package entry). */
export interface RuntimeFileAsset {
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

/** A whole asset directory (cmaps, standard_fonts, ...) copied recursively
 *  into public/ at build time. `dir` resolves relative to the package root of
 *  `resolveFrom` (pdfjs-dist's cmaps/ sits at the package root, not under
 *  build/). Dir entries never enter stableAssetNames — they are fetched at
 *  runtime, not emitted by Rollup (see runtime-assets.ts). */
export interface RuntimeDirAsset {
  resolveFrom: string;
  through?: string;
  /** Directory within the resolved package directory. */
  dir: string;
  /** Destination directory relative to the docs public/ directory. */
  to: string;
}

export type RuntimeAsset = RuntimeFileAsset | RuntimeDirAsset;

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
      // 口径是"文档站收录的包"，即对外发布的业务包——引擎层
      // （xlsx-core）与遮罩层（progress-overlay）已转为仓库私有包、构建期
      // 打进 Excel 业务包 dist，不再独立发布，故不设独立页面。
      // 与 guide/index.md 的仓库结构表一致。
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
          en: "Auto mode picks the best path by rows and estimated cells; business code never changes as data grows.",
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
    // 已发正式版，与导出包同为 stable 口径（版本号动态读取 package.json，勿在此写死纪元）
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
    dir: "word-preview",
    npmName: "@marcusok/word-preview",
    category: "preview",
    version: wordPreviewPkg.version,
    // 首版 0.x：上游分页/TOC 限制与本包嗅探/后处理逻辑需要观察期
    status: "beta",
    zh: true,
    tagline: {
      zh: "docx 只读预览（视觉保真渲染 + 格式嗅探 + 安全加固默认值）",
      en: "Read-only docx preview (visual-faithful rendering, format sniffing, hardened defaults)",
    },
    keywords: ["word", "docx", "preview", "viewer", "docx-preview"],
    highlights: [
      {
        icon: "pen",
        title: { zh: "视觉保真", en: "Visual fidelity" },
        details: {
          zh: "基于 docx-preview 引擎渲染 Word 文档的分页、页眉页脚、脚注尾注、批注与修订；分页遵循 Word 存盘标记，贴近原版观感。",
          en: "Built on the docx-preview engine: pagination, headers/footers, footnotes/endnotes, comments and tracked changes render as Word saved them.",
        },
      },
      {
        icon: "shield",
        title: { zh: "安全加固默认值", en: "Hardened defaults" },
        details: {
          zh: "altChunk（内嵌 HTML 的 iframe）默认关闭；超链接仅放行 http/https/mailto/tel 与锚点/相对形态；渲染前先做格式嗅探与 zip 必需部件校验，坏文件给出可读错误码。",
          en: "altChunks (embedded HTML iframes) are off by default; hyperlink hrefs pass an http/https/mailto/tel allowlist; input is sniffed (magic bytes + required ZIP parts) before rendering, so broken files get readable error codes.",
        },
      },
      {
        icon: "route",
        title: { zh: "多实例隔离", en: "Multi-instance isolation" },
        details: {
          zh: "每个实例生成唯一样式类名，同页渲染多份文档互不串扰（上游固定类名会让后渲染的文档覆盖前者的编号/主题样式）。",
          en: "Each instance gets a unique style class name, so multiple documents on one page never clobber each other's numbering/theme styles (the engine's fixed class name would).",
        },
      },
    ],
    demo: "WordPreviewDemo",
  },
  {
    dir: "pdf-preview",
    npmName: "@marcusok/pdf-preview",
    category: "preview",
    version: pdfPreviewPkg.version,
    // 首版 0.x：懒渲染释放策略与 bundler 场景的资产接入指引需要观察期
    status: "beta",
    zh: true,
    tagline: {
      zh: "PDF 只读预览（pdf.js 连续滚动懒渲染 + 文本层可选中复制）",
      en: "Read-only PDF preview (pdf.js continuous lazy rendering, selectable text layer)",
    },
    keywords: ["pdf", "preview", "viewer", "pdfjs", "text-layer"],
    // worker 零配置（bundler 对锚点字面量自动 emit）；cmaps/standard_fonts
    // 是目录资产、bundler 不复制——经 runtimeAssets 拷入 public/，demo 显式
    // 传 withBase URL（锚点推导在 build 下落在 hashed assets 目录，与
    // public 拷贝对不上，且 cmap 缺失是静默劣化不报错）
    runtimeAssets: [
      {
        resolveFrom: "pdfjs-dist",
        through: "@marcusok/pdf-preview",
        dir: "cmaps",
        to: "cmaps",
      },
      {
        resolveFrom: "pdfjs-dist",
        through: "@marcusok/pdf-preview",
        dir: "standard_fonts",
        to: "standard_fonts",
      },
    ],
    highlights: [
      {
        icon: "zap",
        title: { zh: "大文档流畅滚动", en: "Smooth on large documents" },
        details: {
          zh: "解析在 Web Worker；连续滚动懒渲染——进入视口的页才渲染，远离视口的页自动释放 canvas 显存，重新进入按需重渲。",
          en: "Parsing runs in a Web Worker; continuous scroll with lazy per-page rendering — pages render as they enter the viewport and release their canvas memory when far away.",
        },
      },
      {
        icon: "pen",
        title: { zh: "文本可选中复制", en: "Selectable, copyable text" },
        details: {
          zh: "透明文本层与 canvas 像素精确对齐（pdf.js TextLayer），预览即文档：可选中、复制、检索；不需要时一行选项关闭。",
          en: "A transparent text layer aligns with the canvas pixels (pdf.js TextLayer) — select, copy and search the preview like real text; one option turns it off.",
        },
      },
      {
        icon: "shield",
        title: { zh: "安全加固默认值", en: "Hardened defaults" },
        details: {
          zh: "渲染前先做 %PDF- 魔数嗅探（Adobe 1024 字节窗口口径）；加密文档给出明确错误码；不加载 sandbox 模块，PDF 内嵌 JavaScript 永不执行；注释/超链接层首版不渲染。",
          en: "Input is sniffed for the %PDF- magic (Adobe's 1024-byte window) before the engine runs; encrypted files report a clear code; the sandbox module is never loaded, so embedded PDF JavaScript never executes; the annotation layer is not rendered in v1.",
        },
      },
      {
        icon: "route",
        title: { zh: "资产接入简单", en: "Simple asset wiring" },
        details: {
          zh: "worker/cmaps/standard_fonts 全部随包分发（dist 约 6.4MB 含 sourcemap）；bundler 场景 worker 零配置，cmaps/字体目录可自托管或配 copy 插件（文档给出配方）。",
          en: "The worker, cmaps and standard fonts all ship with the package; workers need zero config under bundlers, while the cmaps/font directories can be self-hosted or wired with a copy plugin (recipes in the docs).",
        },
      },
    ],
    demo: "PdfPreviewDemo",
  },
  {
    dir: "ppt-preview",
    npmName: "@marcusok/ppt-preview",
    category: "preview",
    version: pptPreviewPkg.version,
    // 首版 0.x：上游引擎年轻（0.89.x、维护者集中），保真与降级边界需要观察期
    status: "beta",
    zh: true,
    tagline: {
      zh: "pptx 只读预览（Rust/WASM 引擎连续滚动懒渲染 + 0-based 页导航 + 加密解密）",
      en: "Read-only pptx preview (Rust/WASM engine, continuous lazy rendering, 0-based navigation, password decryption)",
    },
    keywords: ["pptx", "powerpoint", "preview", "viewer", "wasm", "slides"],
    // 单 wasm 随包分发且由 bundler emit（锚点字面量自动资产化），无需
    // runtimeAssets 的 public/ 拷贝（与 pdf-preview 的目录资产场景不同）
    highlights: [
      {
        icon: "zap",
        title: { zh: "大幻灯片文档流畅滚动", en: "Smooth on large decks" },
        details: {
          zh: "解析在 Web Worker（Rust/WASM 引擎）；连续滚动懒渲染——进入视口的幻灯片才渲染，远离视口的页自动释放 canvas 显存。",
          en: "Parsing runs in a Web Worker (a Rust/WASM engine); continuous scroll with lazy per-slide rendering — slides render as they enter the viewport and release canvas memory when far away.",
        },
      },
      {
        icon: "pen",
        title: { zh: "文本可选中复制", en: "Selectable, copyable text" },
        details: {
          zh: "文本 runs 在正式渲染的同一趟收集（无双倍渲染），透明文本层与 canvas 像素对齐——可选中、复制；不需要时一行选项关闭。",
          en: "Text runs are collected in the same render pass (no double rendering); a transparent layer aligns with the canvas — select and copy like real text; one option turns it off.",
        },
      },
      {
        icon: "shield",
        title: { zh: "加密文件可解密", en: "Encrypted files decryptable" },
        details: {
          zh: "支持 PowerPoint「打开密码」（MS-OFFCRYPTO Agile，SHA-512 + AES）；缺密码/密码错误给出可区分的错误码；解密在主线程，约 +1s。",
          en: "PowerPoint open passwords (MS-OFFCRYPTO Agile, SHA-512 + AES) are supported; missing and wrong passwords map to distinguishable error codes; decryption runs on the main thread (~+1s).",
        },
      },
      {
        icon: "route",
        title: { zh: "0-based 页导航", en: "0-based navigation" },
        details: {
          zh: "goToSlide/nextSlide/prevSlide 基于当前最可见页跟踪；隐藏页全部显示并计入 slideCount，页序与原始 deck 一致。",
          en: "goToSlide/nextSlide/prevSlide track the most visible slide; hidden slides are all shown and counted, preserving the original deck order.",
        },
      },
    ],
    demo: "PptPreviewDemo",
  },
];
