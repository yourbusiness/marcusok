import { existsSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { getDemos, groupDemos, groupRepresentative } from "../common/registry";
import type { DemoEntry } from "../common/registry";

const packageRoot = resolve(
  dirname(fileURLToPath(import.meta.url)),
  "..",
  "..",
);

function workspaceDeps(): string[] {
  const json = JSON.parse(
    readFileSync(resolve(packageRoot, "package.json"), "utf8"),
  ) as { dependencies?: Record<string, string> };
  return Object.keys(json.dependencies ?? {})
    .filter((name) => name.startsWith("@marcusok/"))
    .map((name) => name.replace(/^@marcusok\//, ""));
}

describe("demo registration completeness", () => {
  const deps = workspaceDeps();

  it("every @marcusok/* dependency follows the src/index.ts layout convention", () => {
    for (const dep of deps) {
      const pkgDir = resolve(packageRoot, "..", dep);
      const hasEntry =
        existsSync(resolve(pkgDir, "src/index.ts")) ||
        existsSync(resolve(pkgDir, "src/index.tsx"));
      expect(
        hasEntry,
        `@marcusok/${dep} 缺少 src/index.ts(.tsx)，无法获得源码 HMR`,
      ).toBe(true);
    }
  });

  it("every @marcusok/* dependency registers a demo under its own name", () => {
    expect(deps.length).toBeGreaterThan(0);
    // Same discovery mechanism as main.tsx: eager glob executes every demo
    // entry, which is what calls registerDemo().
    const demoEntries = import.meta.glob(
      ["../demos/*/index.ts", "!../demos/_*/**"],
      { eager: true },
    );
    const demoDirs = Object.keys(demoEntries).map(
      (path) => path.match(/\.\.\/demos\/([^/]+)\/index\.ts$/)![1],
    );
    for (const dep of deps) {
      expect(demoDirs, `缺少 src/demos/${dep}/index.ts`).toContain(dep);
    }
    const names = getDemos().map((demo) => demo.name);
    for (const dep of deps) {
      expect(names, `@marcusok/${dep} 未通过 registerDemo() 注册`).toContain(
        dep,
      );
    }
  });

  it("the _template directory is not registered as a demo", () => {
    const names = getDemos().map((demo) => demo.name);
    expect(names).not.toContain("_template");
  });
});

describe("groupDemos", () => {
  const demo = (name: string, group?: string): DemoEntry => ({
    name,
    label: name,
    ...(group !== undefined && { group }),
    load: () => Promise.reject(new Error("test stub")),
  });

  it("groups by demo.group, keeping registration order", () => {
    const { groups, ungrouped } = groupDemos([
      demo("excel-exporter", "excel-exporter"),
      demo("other-pkg"),
      demo("excel-exporter-styles", "excel-exporter"),
    ]);
    expect(groups).toEqual([
      [
        "excel-exporter",
        [
          expect.objectContaining({ name: "excel-exporter" }),
          expect.objectContaining({ name: "excel-exporter-styles" }),
        ],
      ],
    ]);
    expect(ungrouped.map((d) => d.name)).toEqual(["other-pkg"]);
  });

  it("keeps several groups separate and in first-seen order", () => {
    const { groups, ungrouped } = groupDemos([
      demo("pkg-a", "pkg-a"),
      demo("solo"),
      demo("pkg-b", "pkg-b"),
      demo("pkg-a-extra", "pkg-a"),
      demo("pkg-b-extra", "pkg-b"),
    ]);
    expect(groups.map(([name, members]) => [name, members.length])).toEqual([
      ["pkg-a", 2],
      ["pkg-b", 2],
    ]);
    expect(ungrouped.map((d) => d.name)).toEqual(["solo"]);
  });

  it("groupRepresentative prefers the demo named after the group", () => {
    const children = [
      demo("excel-exporter-styles", "excel-exporter"),
      demo("excel-exporter", "excel-exporter"),
    ];
    expect(groupRepresentative("excel-exporter", children).name).toBe(
      "excel-exporter",
    );
    // 无同名 demo 时回退组内第一个
    expect(groupRepresentative("excel-exporter", [children[0]!]).name).toBe(
      "excel-exporter-styles",
    );
  });
});
