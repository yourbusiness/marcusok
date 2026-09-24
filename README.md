# marcusok

A pnpm + Turborepo frontend-infrastructure monorepo providing shared capability packages for multiple admin applications. Packages are organized in two categories: **Export** (turning data into downloadable documents — an Excel export engine built on [modern-xlsx](https://github.com/ABCrimson/modern-xlsx) (Rust + WASM) is available today, other document formats may follow) and **Document preview** (rendering documents in the browser, planned).

## Packages

| Package                                                 | Category | Description                                                                                                      |
| ------------------------------------------------------- | -------- | ---------------------------------------------------------------------------------------------------------------- |
| [`@marcusok/excel-exporter`](./packages/excel-exporter) | Export   | Excel export core library (WASM-driven, styled, streaming writes, Worker multithreading, table/ECharts adapters) |

## Quick Start

```bash
pnpm install     # install dependencies
pnpm build       # build all packages
pnpm test        # run all tests
pnpm lint        # ESLint
pnpm typecheck   # TypeScript type checking
```

Environment: Node >= 22.12 (the floor Vite 8 requires; the published `@marcusok/excel-exporter` package itself only needs Node >= 22), pnpm >= 9. `.nvmrc` pins Node 22. The modern-xlsx@1.2.0 devDependency (bundled into the published package at build time) declares `engines.node >= 24`, but its WASM core targets browsers; the repo is fully green on Node 22, and `.npmrc` sets `engine-strict=false` to allow this. Consumers are unaffected — the published package has zero runtime dependencies.

## Tooling

| Area               | Choice                           | Notes                                                                                                                                     |
| ------------------ | -------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| Package manager    | pnpm workspace                   | Hard links save disk, `workspace:*` protocol, strict dependency isolation                                                                 |
| Build orchestrator | Turborepo                        | Parallel builds + local cache; `^build` automatically orders package dependencies                                                         |
| Package bundler    | tsup                             | esbuild-driven; TS → ESM + DTS in one pass; ESM-only across the board                                                                     |
| Language           | TypeScript 5.x                   | `moduleResolution: bundler`; `lib` includes both DOM and WebWorker                                                                        |
| Code style         | ESLint 9 + Prettier              | flat config; `no-floating-promises` prevents missed awaits                                                                                |
| Commit convention  | Husky + lint-staged + commitlint | Conventional Commits, powering Changesets-generated changelogs                                                                            |
| Versioning/release | Changesets                       | Independent per-package releases, auto-generated changelogs, prerelease support                                                           |
| Testing            | Vitest                           | Native ESM, WASM-friendly                                                                                                                 |
| CI/CD              | GitHub Actions                   | `ci.yml` (commitlint → format:check → lint → typecheck → test → build) + `release.yml` (auto publish) + `deploy.yml` (docs site to Pages) |

## Directory Layout

```
marcusok/
├── apps/                       # Applications (scales horizontally)
│   └── docs/                   # VitePress public docs site (English default + Chinese, GitHub Pages)
├── packages/                   # Shared packages (scales horizontally)
│   ├── excel-exporter/         # Export category: Excel export
│   │   ├── src/                # Source (incl. workers/ entry) and __tests__/
│   │   └── dist/               # tsup build output
│   └── play/                   # Local integration sandbox (React 19 + antd 6, private package)
├── docs/                       # Design documents (Chinese)
│   ├── excel-export-design.md  # Excel export core design doc (~230k chars, the main one)
│   └── release-*.md / ci-*.md  # Release & CI walkthroughs, debug notes, docs-site plan
├── scripts/                    # Repo-level scripts (dev.mjs unified dev launcher)
├── .changeset/                 # Changesets config
├── .github/workflows/          # CI/CD
├── turbo.json                  # Turborepo task orchestration
├── tsconfig.base.json          # Shared TypeScript baseline
├── pnpm-workspace.yaml
├── eslint.config.mjs           # ESLint flat config
└── package.json
```

## Adding a New Package

1. Create it under `packages/<name>/` and reference internal deps with `workspace:*`.
2. The package's `package.json` declares `type: "module"`; build with tsup, test with Vitest.
3. Reuse the repo-root `tsconfig.base.json` and `eslint.config.mjs` for cross-package shared config (there is no `packages/_shared/` yet; extract one when a real need arises).
4. Turborepo's `^build` dependency graph handles build order automatically; new packages require no CI/CD changes.
5. Changesets releases each package independently — nothing blocks anything else.
6. Declare the package's category (`export` / `preview`) in the docs registry (`apps/docs/.vitepress/registry.ts`) so the docs site groups it correctly; categories live only in that registry, not in the directory layout.

## Release Process

Two-phase by design — `release.yml` runs on every push to `main`:

```bash
pnpm changeset                # create a changeset, pick affected packages and semver type
# commit .changeset/*.md → merge to main → release.yml opens a "chore: release packages" PR:
#   changeset version         bump versions + update CHANGELOGs (staged in that PR, NOT published)
# merge that PR → release.yml runs again and, with .changeset/ now empty, executes:
#   pnpm release              format:check + lint + typecheck + test + build, then changeset publish
```

Merging the version PR is what publishes to npm — it is a human decision, never automatic.

Prerelease:

```bash
pnpm changeset pre enter next  # enter next prerelease mode
pnpm changeset version         # → 0.1.3-next.0
pnpm changeset publish         # publish with the next dist-tag
```

> The prerelease flow runs locally and is the one intentional exception: it
> publishes directly and therefore skips the quality gate that the automated
> release flow enforces. Run `pnpm lint && pnpm typecheck && RUN_PERF=0 pnpm test && pnpm build`
> yourself before publishing, and leave `changeset pre exit` to return to the
> normal two-phase (version PR → merge to publish) flow.

## Reference Docs

- [`docs/excel-export-design.md`](./docs/excel-export-design.md) — Excel export core design doc
- [`docs/release-guide.md`](./docs/release-guide.md) — Release guide
- [`docs/release-publish-logic.md`](./docs/release-publish-logic.md) — Release logic in depth
- [`docs/release-workflow-analysis.md`](./docs/release-workflow-analysis.md) — Release workflow analysis
- [`docs/ci-workflow-analysis.md`](./docs/ci-workflow-analysis.md) — CI workflow analysis
- [`docs/changeset-walkthrough.md`](./docs/changeset-walkthrough.md) — Changesets usage notes
- [`docs/vitepress-docs-plan.md`](./docs/vitepress-docs-plan.md) — Docs site planning
- [`docs/debug.md`](./docs/debug.md) — Debugging guide

## License

MIT
