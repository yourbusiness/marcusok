# play

The monorepo's local integration sandbox, built with **React 19 + TypeScript + Vite 8 + Ant Design 6**. Each package under test gets its own directory under `src/demos/<pkg>/`; demo implementations are loaded on demand and the home page only loads lightweight metadata.

## Tech Stack

| Dependency           | Version notes                                                          |
| -------------------- | ---------------------------------------------------------------------- |
| react / react-dom    | 19.2.x                                                                 |
| antd                 | 6.5.x (native React 19 support)                                        |
| vite                 | 8.2.x (Rolldown core)                                                  |
| @vitejs/plugin-react | 6.x                                                                    |
| vitest               | 4.1.x (peer supports vite 8)                                           |
| typescript           | 5.9.x (typescript-eslint 8.65 constrains < 6.1; not upgrading to TS 7) |

## Usage

```bash
# Recommended: start from the repo root (unified dev launcher scripts/dev.mjs)
pnpm dev

# play only: build upstream packages first, then start vite (no tsup --watch or vitepress)
pnpm dev:play
```

The difference:

- `pnpm dev`: runs `node scripts/dev.mjs`, which first runs `turbo run build` once
  (building `excel-exporter` and friends into `dist/`, equivalent to turbo dev's
  `dependsOn: ["^build"]`), then starts all three dev services in parallel:
  excel-exporter's `tsup --watch`, play's `vite` (5173), and docs'
  `vitepress dev` (5174).
- `pnpm dev:play`: runs `node scripts/dev.mjs play` — same upstream build, then
  only play's vite.

> The dev launcher (`scripts/dev.mjs`) has replaced `turbo run dev`: on Windows,
> the turbo → pnpm.CMD → cmd.exe → node wrapper chain breaks Ctrl+C signal
> delivery, leaving orphaned child processes holding ports; the launcher spawns
> each process directly from node and cleans up the process tree on exit.

The Vite dev server is pinned to `http://localhost:5173` (`strictPort`: if the
port is taken it fails loudly instead of silently moving to 5174/5175; use
`vite --port <n>` if you really need another port).

Quality gates:

```bash
pnpm --filter @marcusok/play typecheck
pnpm --filter @marcusok/play lint
pnpm --filter @marcusok/play test
```

> play is a `private` package and does not participate in changeset publishing;
> `build` (`vite build`) serves only as a buildability check (run by CI and the
> root `pnpm build`; the output `dist/` is not committed).

## Layout & Interaction

- Collapsible dark sidebar (brand area + demo list) plus a top bar (current demo
  name + light/dark theme toggle).
- The home page shows demos as a card grid (label + description, hover feedback).
- Theming is centralized in `src/app/theme.ts`: primary color, border radius,
  light/dark algorithms; the selection persists to localStorage.
- Routing is a lightweight hash router (`#/excel-exporter`): refresh-safe and
  shareable URLs; with only an "overview / detail" structure, react-router is
  unnecessary.

## Mock Data

`src/mock/rows.ts` provides a deterministic mock generator (mulberry32 PRNG with
a fixed seed):

- Tiers: `100 / 1k / 10k / 50k / 100k / 200k`, exposed as the `DATASET_PRESETS`
  constant.
- Default tier: `10,000` rows (`DEFAULT_ROWS`; change this one constant to
  adjust the default).
- Fields mix strings, numbers, dates and enums to exercise excel-exporter's
  `format` / `numFormat` handling.

With the same seed and tier the data is identical every time, so repeated
exports and cross-run performance comparisons are not distorted by data
randomness.

## Adding a Package Sandbox

1. Copy `src/demos/_template/` → `src/demos/<your-pkg>/`
2. Declare `"@marcusok/<your-pkg>": "workspace:*"` in play's `package.json`
   dependencies, then run `pnpm install`
3. Uncomment the template in `index.ts`, fill in `name` / `label` /
   `description`, and dynamically import the implementation in `load()`
   (returning `{ default: React component }`)
4. **Restart the dev server**: `import.meta.glob` is expanded statically at
   startup, so demo directories added at runtime are not discovered (adding a
   demo or renaming one both require a restart)

### Package layout conventions (enforced by tests)

`pnpm --filter @marcusok/play test` validates the following rules; a
non-compliant package fails the suite outright:

- Every package under test must provide `src/index.ts` (or `src/index.tsx`) as
  its main entry, otherwise it cannot get source-alias HMR. Vite prints a
  warning for non-compliant dependency packages at startup and falls back to
  dist resolution.
- Every `@marcusok/*` package declared in dependencies must have
  `src/demos/<pkg>/index.ts` registering itself via
  `registerDemo({ name: "<pkg>", ... })`.

### Demo lifecycle

Each demo's `index.ts` entry only registers lightweight metadata (`name` /
`label` / `description` / `load`); the actual UI lives in a separate
`*.demo.tsx`, dynamically imported by `load()` on demand:

```ts
// index.ts — metadata only; do not statically import heavy dependencies here
registerDemo({
  name: "your-pkg",
  label: "your-pkg · summary",
  description: "One sentence on what this demo shows.",
  async load() {
    return import("./your-pkg.demo.js");
  },
});
```

The module returned by `load()` must default-export a React component, rendered
inside the Suspense boundary in `App.tsx`; resource cleanup (timers, fetch,
WebSocket, etc.) happens in the component's `useEffect` cleanup, which runs
automatically on navigation — no hand-written `destroy()` mechanism is needed.

## Module Resolution Rules (vite.config.ts)

- Main entry `@marcusok/<pkg>` → `src/index.ts(.tsx)` source (HMR-friendly).
- Subpath `@marcusok/<pkg>/<sub>` → tried in order: `src/<sub>.ts(.tsx)` /
  `src/<sub>/index.*` source → mapped back to source via the package's `exports`
  (nested conditions supported, e.g. `{"import": {"types": ..., "default": ...}}`;
  source resolution works even when the export key differs from the source file
  name) → `dist/<sub>` build output. So subpaths get HMR too; subpaths that use
  `dist/` artifacts (like workers) require an upstream build.
- Third-party subpaths not exposed via `exports` (e.g. `modern-xlsx/wasm/*`):
  add a `{ pkg, dir, excludeFromOptimizeDeps }` entry to `externalOverrides` in
  `vite.config.ts` and the resolver rewrites it to the physical file.
- The pure-function resolution/alias logic lives in
  `src/vite/workspace-resolver.ts` with unit-test coverage
  (`src/__tests__/workspace-resolver.test.ts`).

## Worker / WASM Assets

If your package uses runtime assets like Workers or WASM (not resolvable via
plain imports), import the asset path with Vite's `?url` suffix and pass it to
the package's configure function. See the `excel-exporter` demo:

```ts
import workerUrl from "@marcusok/your-pkg/dist/your.worker.js?url";
import wasmUrl from "your-wasm-dep/your.wasm?url";

yourPkg.configure({ workerUrl, wasmUrl });
```

Note that `@marcusok/<pkg>/dist/*.worker.js` is build output, so upstream
packages must be built (both `pnpm dev` and `pnpm dev:play` do this
automatically). Changes to dist output do not trigger HMR (the source alias only
covers the main entry and subpath sources) — after changing a worker or similar
artifact, rebuild the upstream package and refresh the page manually. Without
worker/wasm configured, a package may silently degrade to a fallback path (e.g.
excel-exporter falls back to SheetJS, losing styles).

## Toolchain

- The pnpm version is pinned by the root `package.json`'s `packageManager`
  field (`pnpm@9.12.0`); run it via corepack (or equivalent) so different pnpm
  versions don't rewrite the lockfile and produce huge diffs.
- play uses Vite 8, which requires Node >= 22.12 (declared in play's
  `package.json` `engines`); the local Node 22.22.2 satisfies this.
- vitest 4.1.x aligns with vite 8 (peer supports `^6 || ^7 || ^8`); the repo
  carries a single vite major version.

## HMR

After editing demo source, `@vitejs/plugin-react` provides React Fast Refresh —
the current demo hot-updates in place without a full page reload, and the
current hash route is preserved. When a demo component unmounts, React
automatically runs its `useEffect` cleanup (equivalent to the hand-written
`destroy()` in the old HMR setup).
