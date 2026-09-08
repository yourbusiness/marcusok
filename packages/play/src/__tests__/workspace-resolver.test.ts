import { describe, expect, it } from "vitest";
import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { dirname, join, resolve } from "node:path";
import { tmpdir } from "node:os";
import { fileURLToPath } from "node:url";
import {
  isFile,
  matchExportTarget,
  pickExportTarget,
  readExportsMap,
  resolvePkgDir,
  srcCandidatesFromDistTarget,
} from "../vite/workspace-resolver";

const playDir = resolve(dirname(fileURLToPath(import.meta.url)), "..", "..");
const exporterDir = resolve(playDir, "..", "excel-exporter");

describe("pickExportTarget", () => {
  it("passes through a plain string target", () => {
    expect(pickExportTarget("./dist/index.js")).toBe("./dist/index.js");
  });

  it("prefers import over default/require/types", () => {
    const value = {
      types: "./dist/index.d.ts",
      import: "./dist/index.js",
      require: "./dist/index.cjs",
      default: "./dist/index.js",
    };
    expect(pickExportTarget(value)).toBe("./dist/index.js");
  });

  it("unwraps nested condition objects (dual-package exports)", () => {
    const value = {
      import: { types: "./dist/index.d.ts", default: "./dist/index.js" },
      require: "./dist/index.cjs",
    };
    expect(pickExportTarget(value)).toBe("./dist/index.js");
  });

  it("falls back to the next condition when a branch has no runtime target", () => {
    const value = { import: {}, default: "./dist/index.js" };
    expect(pickExportTarget(value)).toBe("./dist/index.js");
  });

  it("returns empty string for unsupported shapes", () => {
    expect(pickExportTarget(null)).toBe("");
    expect(pickExportTarget(undefined)).toBe("");
    expect(pickExportTarget([])).toBe("");
    expect(pickExportTarget(42)).toBe("");
  });
});

describe("matchExportTarget", () => {
  const map = new Map<string, string>([
    ["./styles", "./dist/style-presets.js"],
    ["./worker-utils", "./dist/worker-utils.js"],
    ["./dist/*.js", "./dist/*.js"],
  ]);

  it("matches exact keys", () => {
    expect(matchExportTarget(map, "styles")).toBe("./dist/style-presets.js");
  });

  it("matches single-star wildcard keys", () => {
    expect(matchExportTarget(map, "dist/export.worker.js")).toBe(
      "./dist/export.worker.js",
    );
  });

  it("returns empty string when nothing matches", () => {
    expect(matchExportTarget(map, "nope")).toBe("");
  });
});

describe("srcCandidatesFromDistTarget", () => {
  it("maps a dist file back to plausible source candidates", () => {
    expect(srcCandidatesFromDistTarget("./dist/style-presets.js")).toEqual([
      "style-presets.ts",
      "style-presets.tsx",
      "style-presets/index.ts",
      "style-presets/index.tsx",
    ]);
  });

  it("handles nested dist paths", () => {
    expect(srcCandidatesFromDistTarget("./dist/workers/export.js")).toEqual([
      "workers/export.ts",
      "workers/export.tsx",
      "workers/export/index.ts",
      "workers/export/index.tsx",
    ]);
  });

  it("returns empty for non-dist targets", () => {
    expect(srcCandidatesFromDistTarget("./src/index.ts")).toEqual([]);
  });
});

describe("isFile", () => {
  it("returns true for regular files, false for directories and missing paths", () => {
    expect(isFile(resolve(playDir, "package.json"))).toBe(true);
    expect(isFile(resolve(playDir, "src"))).toBe(false);
    expect(isFile(resolve(playDir, "no-such-file.xyz"))).toBe(false);
  });
});

describe("resolvePkgDir", () => {
  it("locates an installed third-party package root through pnpm symlinks", () => {
    // vitest is a direct devDependency of play, so it is resolvable from
    // play's node_modules; the root must be the dir whose package.json name
    // matches (not an intermediate dist/ directory).
    const dir = resolvePkgDir("vitest", playDir);
    expect(dir).not.toBe("");
    const pkg = JSON.parse(
      readFileSync(resolve(dir, "package.json"), "utf8"),
    ) as { name?: string };
    expect(pkg.name).toBe("vitest");
  });

  it("locates a workspace dependency at its monorepo source directory", () => {
    // Fresh checkouts run this test BEFORE the workspace package's first
    // build (turbo's test task has no ^build dependency), so "exports"."."
    // may point at a dist/index.js that does not exist yet — the fixture
    // below covers that fallback branch; with dist present the main-entry
    // branch resolves. Either way the root must be the source directory.
    const dir = resolvePkgDir("@marcusok/excel-exporter", playDir);
    expect(dir).toBe(exporterDir);
  });

  it("falls back to <pkg>/package.json when the main entry artifact is missing", () => {
    // Hermetic fixture: a node_modules package whose exports map points its
    // main entry at a nonexistent dist file (a fresh, unbuilt checkout) but
    // exposes "./package.json" — the state where the previous implementation
    // threw inside require.resolve and returned "".
    const base = mkdtempSync(join(tmpdir(), "resolver-fallback-"));
    try {
      const pkgDir = join(base, "node_modules", "stub-pkg");
      mkdirSync(pkgDir, { recursive: true });
      writeFileSync(
        join(pkgDir, "package.json"),
        JSON.stringify({
          name: "stub-pkg",
          exports: {
            ".": "./dist/never-built.js",
            "./package.json": "./package.json",
          },
        }),
      );
      expect(resolvePkgDir("stub-pkg", base)).toBe(pkgDir);
    } finally {
      rmSync(base, { recursive: true, force: true });
    }
  });

  it("returns an empty string for unresolvable packages", () => {
    expect(resolvePkgDir("definitely-not-a-real-package-xyz", playDir)).toBe(
      "",
    );
  });
});

describe("readExportsMap", () => {
  it("flattens the real excel-exporter exports map via the import condition", () => {
    const map = readExportsMap(exporterDir);
    expect(map.get(".")).toBe("./dist/index.js");
    expect(map.get("./styles")).toBe("./dist/style-presets.js");
    expect(map.get("./worker-utils")).toBe("./dist/worker-utils.js");
    expect(map.get("./dist/export.worker.js")).toBe("./dist/export.worker.js");
    expect(map.get("./package.json")).toBe("./package.json");
  });

  it("returns an empty map for a package without an exports field", () => {
    // play itself is a private app with no exports map.
    expect(readExportsMap(playDir).size).toBe(0);
  });
});
