import docsAppPkg from "../package.json" with { type: "json" };
import { packages } from "./registry";

const RESERVED_STAT_KEYS = new Set(["packages"]);

/** Loose semver guard: x.y.z with optional -prerelease / +build. */
const SEMVER_RE = /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$/;
const SAFE_DIR_RE = /^[a-z0-9][a-z0-9-]*$/;

/**
 * Fail fast (at build time) on registry mistakes that would break the site.
 *
 * Node-only: this module reads the docs app's own package.json to verify that
 * every registered package is declared as a dependency. It is imported solely
 * by config.ts (the VitePress config, a Node context) and must NEVER enter the
 * browser bundle — the theme components import only the pure `registry.ts`.
 */
export function validateRegistry(): void {
  const seenDirs = new Set<string>();
  const seenNames = new Set<string>();
  const seenStatKeys = new Set<string>(RESERVED_STAT_KEYS);
  const seenAssetTargets = new Set<string>();
  const declaredDeps = docsAppPkg.dependencies ?? {};
  for (const p of packages) {
    if (seenDirs.has(p.dir)) {
      throw new Error(`[registry] duplicate dir "${p.dir}"`);
    }
    seenDirs.add(p.dir);
    if (seenNames.has(p.npmName)) {
      throw new Error(`[registry] duplicate npmName "${p.npmName}"`);
    }
    seenNames.add(p.npmName);
    if (!SAFE_DIR_RE.test(p.dir)) {
      throw new Error(
        `[registry] "${p.npmName}": dir "${p.dir}" must match ${SAFE_DIR_RE}`,
      );
    }
    if (!Object.hasOwn(declaredDeps, p.npmName)) {
      throw new Error(
        `[registry] "${p.npmName}" is not declared in apps/docs/package.json dependencies. ` +
          'Add it (e.g. "workspace:*") before registering the docs entry.',
      );
    }
    if (!SEMVER_RE.test(p.version)) {
      throw new Error(
        `[registry] "${p.npmName}": version "${p.version}" is not a valid semver. ` +
          "Read it from the package's own package.json (single source of truth).",
      );
    }
    for (const s of p.homeStats ?? []) {
      if (seenStatKeys.has(s.key)) {
        throw new Error(
          `[registry] duplicate homeStat key "${s.key}" (${p.npmName})`,
        );
      }
      seenStatKeys.add(s.key);
      if (!Number.isFinite(s.value) || s.decimals < 0 || s.decimals > 3) {
        throw new Error(
          `[registry] "${p.npmName}": homeStat "${s.key}" has invalid value/decimals`,
        );
      }
    }
    const sectionIds = new Set<string>();
    for (const s of p.sections ?? []) {
      if (sectionIds.has(s.id)) {
        throw new Error(
          `[registry] duplicate section id "${s.id}" (${p.npmName})`,
        );
      }
      sectionIds.add(s.id);
    }
    // runtimeAssets copy by plain copyFileSync in registration order: two
    // packages targeting the same destination would silently overwrite each
    // other's asset. Reject the collision up front.
    for (const a of p.runtimeAssets ?? []) {
      if (seenAssetTargets.has(a.to)) {
        throw new Error(
          `[registry] duplicate runtimeAssets target "${a.to}" (${p.npmName}): ` +
            "a later copy would silently overwrite an earlier package's asset",
        );
      }
      seenAssetTargets.add(a.to);
    }
    // Benchmark data must line up with its series definitions: BenchmarkChart
    // reads values by series key with a `?? 0` fallback, so a typo'd key would
    // silently render a zero-height bar instead of failing the build.
    for (const b of p.benchmarks ?? []) {
      if (b.series.length === 0) {
        throw new Error(
          `[registry] "${p.npmName}": benchmark must declare at least one series`,
        );
      }
      const seriesKeys = new Set<string>();
      for (const s of b.series) {
        if (seriesKeys.has(s.key)) {
          throw new Error(
            `[registry] duplicate benchmark series key "${s.key}" (${p.npmName})`,
          );
        }
        seriesKeys.add(s.key);
      }
      for (const bar of b.data) {
        for (const [k, v] of Object.entries(bar.values)) {
          if (!seriesKeys.has(k)) {
            throw new Error(
              `[registry] "${p.npmName}": benchmark bar "${bar.label}" value key "${k}" ` +
                `matches no series key [${[...seriesKeys].join(", ")}]`,
            );
          }
          if (!Number.isFinite(v) || v <= 0) {
            throw new Error(
              `[registry] "${p.npmName}": benchmark bar "${bar.label}" series "${k}" ` +
                `has non-positive/non-finite value ${v} (the chart uses log scale)`,
            );
          }
        }
      }
    }
  }
}
