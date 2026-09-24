# Ecosystem

`marcusok` is a pnpm + Turborepo frontend monorepo providing shared capability packages for multiple admin applications. Every package is versioned and published independently via Changesets, wired together with `workspace:*` during development.

## Current packages

| Package                                                 | Category | Status | Description                                                                                                  |
| ------------------------------------------------------- | -------- | ------ | ------------------------------------------------------------------------------------------------------------ |
| [`@marcusok/excel-exporter`](/packages/excel-exporter/) | Export   | stable | Excel export engine: modern-xlsx + Fast stream, full styling, Worker threading, fast writes, stream fallback |

## Engineering conventions

- Package manager: pnpm workspace (`pnpm >= 9`)
- Build orchestration: Turborepo (`^build` resolves package order automatically)
- Package build: tsup (TS → ESM + DTS, ESM-only)
- Language: TypeScript 5.x (`moduleResolution: bundler`)
- Tests: Vitest; linting: ESLint 9 + Prettier
- Versioning/publishing: Changesets (independent versions, changelog, prerelease support)
- CI/CD: GitHub Actions (`ci.yml` checks, `release.yml` publishes to npm, `deploy.yml` ships this docs site to GitHub Pages)

## Roadmap

The ecosystem grows on demand and is organized around two package categories:

- **Export** — turn application data into downloadable documents. Excel export ([`@marcusok/excel-exporter`](/packages/excel-exporter/)) is available today; other document formats (e.g. PDF) may follow.
- **Document preview** — preview documents in the browser. Nothing shipped in this category yet.

When a new package ships, register it in `apps/docs/.vitepress/registry.ts` with its category and add its docs; it automatically appears in the navigation, sidebar and home cards.
