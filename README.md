# marcus-monorepo

A pnpm + Turborepo frontend-infrastructure monorepo providing shared capability packages for multiple admin applications. The first shared package is an Excel export engine built on [modern-xlsx](https://github.com/ABCrimson/modern-xlsx) (Rust + WASM); PDF export, file upload, virtual table rendering and more will follow.

## Packages

| Package                                                 | Description                                                                                                      |
| ------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| [`@marcusok/excel-exporter`](./packages/excel-exporter) | Excel export core library (WASM-driven, styled, streaming writes, Worker multithreading, table/ECharts adapters) |

## Quick Start

```bash
pnpm install     # install dependencies
pnpm build       # build all packages
pnpm test        # run all tests
pnpm lint        # ESLint
pnpm typecheck   # TypeScript type checking
```

Environment: Node >= 22, pnpm >= 9. `.nvmrc` pins Node 22. The modern-xlsx@1.2.0 dependency declares `engines.node >= 24`, but its WASM core targets browsers; the repo is fully green on Node 22, and `.npmrc` sets `engine-strict=false` to allow this.

## Tooling

| Area               | Choice                           | Notes                                                                                                         |
| ------------------ | -------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| Package manager    | pnpm workspace                   | Hard links save disk, `workspace:*` protocol, strict dependency isolation                                     |
| Build orchestrator | Turborepo                        | Parallel builds + local cache; `^build` automatically orders package dependencies                             |
| Package bundler    | tsup                             | esbuild-driven; TS → ESM + DTS in one pass; ESM-only across the board                                         |
| Language           | TypeScript 5.x                   | `moduleResolution: bundler`; `lib` includes both DOM and WebWorker                                            |
| Code style         | ESLint 9 + Prettier              | flat config; `no-floating-promises` prevents missed awaits                                                    |
| Commit convention  | Husky + lint-staged + commitlint | Conventional Commits, powering Changesets-generated changelogs                                                |
| Versioning/release | Changesets                       | Independent per-package releases, auto-generated changelogs, prerelease support                               |
| Testing            | Vitest                           | Native ESM, WASM-friendly                                                                                     |
| CI/CD              | GitHub Actions                   | `ci.yml` (lint → typecheck → test → build) + `release.yml` (auto publish) + `deploy.yml` (docs site to Pages) |

## Directory Layout

```
marcus-monorepo/
├── apps/                       # Applications (scales horizontally)
│   └── docs/                   # VitePress public docs site (English default + Chinese, GitHub Pages)
├── packages/                   # Shared packages (scales horizontally)
│   ├── excel-exporter/         # Current: Excel export
│   │   ├── src/                # Source (incl. workers/ entry) and __tests__/
│   │   └── dist/               # tsup build output
│   └── play/                   # Local integration sandbox (React 19 + antd 6, private package)
├── docs/                       # Design documents
│   └── excel-export-design.md  # Excel export core design doc (~170k chars)
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

## Release Process

```bash
pnpm changeset                # create a changeset, pick affected packages and semver type
# commit .changeset/*.md → merge to main → release.yml automatically runs:
#   changeset version         bump versions + update CHANGELOGs
#   changeset publish         publish to npm
```

Prerelease:

```bash
pnpm changeset pre enter next  # enter next prerelease mode
pnpm changeset version         # → 0.1.3-next.0
pnpm changeset publish         # publish with the next dist-tag
```

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
