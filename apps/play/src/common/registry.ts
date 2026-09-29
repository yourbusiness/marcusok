import type { ComponentType } from "react";

/**
 * 包大类，与文档站 registry 的分类口径一致（导出 / 文档预览）。
 * category 管包所属大类（侧边栏与首页按它分区），与 group（包内 demo
 * 聚合为子菜单）正交。
 */
export type DemoCategory = "export" | "preview";

/** 全部分类与菜单文案（含展示顺序），新增分类时在此登记。 */
export const DEMO_CATEGORIES: { id: DemoCategory; label: string }[] = [
  { id: "export", label: "导出" },
  { id: "preview", label: "文档预览" },
];

export interface DemoEntry {
  name: string;
  /** 包所属大类：侧边栏与首页按它分区（见 DEMO_CATEGORIES）。 */
  category: DemoCategory;
  label: string;
  /** 一句话说明这个 demo 演示什么，展示在首页卡片与详情页头部。 */
  description?: string;
  /**
   * 侧边栏父菜单名：同一 group 的 demo 聚合为一个子菜单，父项本身不可
   * 点击（路由仍以 name 定位，分组只影响菜单呈现）。不设置则为一级菜单。
   */
  group?: string;
  /**
   * 菜单里的显示名；分组场景下 label 通常带包名前缀（如
   * "excel-exporter — …"），不适合直接当子菜单项文案，用它给菜单一个
   * 简短名字。缺省回退 label。
   */
  menuLabel?: string;
  /**
   * Lazy-load the demo implementation. The entry module (index.ts) must
   * stay lightweight — metadata only — so the home page never pulls heavy
   * dependencies. Put the actual UI in a separate *.demo.tsx file and
   * dynamic-import it here; the returned module's default export is a React
   * component rendered by App.tsx's state-driven LazyDemo (Spin placeholder
   * while loading, then the component — no lazy()/Suspense, per React
   * Compiler's static-components constraint).
   */
  load: () => Promise<{ default: ComponentType }>;
}

// Map keyed by name so HMR re-execution of a demo module (which re-calls
// registerDemo) overwrites instead of appending — no duplicate nav entries.
const demos = new Map<string, DemoEntry>();

export function registerDemo(entry: DemoEntry): void {
  const existing = demos.get(entry.name);
  if (existing) {
    console.warn(
      `[play] demo "${entry.name}" 重复注册（旧 label: "${existing.label}"，新 label: "${entry.label}"）。请检查 src/demos/ 下是否有重名目录。`,
    );
  }
  demos.set(entry.name, entry);
}

export function getDemos(): readonly DemoEntry[] {
  return [...demos.values()];
}

export interface DemoCategorySection {
  category: DemoCategory;
  label: string;
  demos: DemoEntry[];
}

/**
 * 按大类分区（保持 DEMO_CATEGORIES 顺序、区内保持注册顺序）。
 * 没有包的大类不产出分区——空分类不渲染标题。
 */
export function demosByCategory(
  entries: readonly DemoEntry[],
): DemoCategorySection[] {
  return DEMO_CATEGORIES.flatMap((c) => {
    const matched = entries.filter((d) => d.category === c.id);
    return matched.length > 0
      ? [{ category: c.id, label: c.label, demos: matched }]
      : [];
  });
}

/**
 * 按 group 聚合 demo（保持注册顺序）：侧边栏菜单与首页卡片共用这套分组。
 * 每个分组的代表 demo 是 name === group 的那个（仓库约定：包级 demo 以包名
 * 命名，registry 测试强制其存在）——分组卡片的文案与点击跳转都落在它身上；
 * 没有同名 demo 时回退到组内第一个。
 */
export function groupDemos(entries: readonly DemoEntry[]): {
  groups: [string, DemoEntry[]][];
  ungrouped: DemoEntry[];
} {
  const groups = new Map<string, DemoEntry[]>();
  const ungrouped: DemoEntry[] = [];
  for (const demo of entries) {
    if (demo.group) {
      const list = groups.get(demo.group) ?? [];
      list.push(demo);
      groups.set(demo.group, list);
    } else {
      ungrouped.push(demo);
    }
  }
  return { groups: [...groups.entries()], ungrouped };
}

/**
 * 分组的代表 demo：name === group 者优先，否则组内第一个。分组由 groupDemos
 * 产出、必非空，因此返回值恒有定义。
 */
export function groupRepresentative(
  group: string,
  members: DemoEntry[],
): DemoEntry {
  return members.find((d) => d.name === group) ?? members[0]!;
}
