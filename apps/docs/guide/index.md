# Ecosystem

`marcusok` is a pnpm + Turborepo frontend monorepo providing shared capability packages for multiple admin applications. Every package is versioned and published independently via Changesets, wired together with `workspace:*` during development.

## Current packages

| Package                                                 | Category | Status | Description                                                                                                  |
| ------------------------------------------------------- | -------- | ------ | ------------------------------------------------------------------------------------------------------------ |
| [`@marcusok/excel-exporter`](/packages/excel-exporter/) | Export   | stable | Excel export engine: modern-xlsx + Fast stream, full styling, Worker threading, fast writes, stream fallback |
| [`@marcusok/excel-preview`](/packages/excel-preview/)   | Preview  | stable | Read-only xlsx preview: worker parsing + virtual scrolling, styles/merges/freeze/number formats restored     |

Two more capability layers exist in the repo as **private internal packages** — the modern-xlsx engine layer (`xlsx-core`: WASM loading, asset distribution, stable re-export surface) and the progress overlay UI (`progress-overlay`, behind the exporter's `overlay` option). They are bundled into the published packages' `dist` at build time and are intentionally not on npm: installing a business package brings everything it needs.

## Engineering conventions

- Package manager: pnpm workspace (`pnpm >= 9`)
- Build orchestration: Turborepo (`^build` resolves package order automatically)
- Package build: tsup (TS → ESM + DTS, ESM-only)
- Language: TypeScript 5.x (`moduleResolution: bundler`)
- Tests: Vitest; linting: ESLint 9 + Prettier
- Versioning/publishing: Changesets (independent versions, changelog, prerelease support)
- CI/CD: GitHub Actions (`ci.yml` checks, `release.yml` publishes to npm, `deploy.yml` ships this docs site to GitHub Pages)

## Roadmap

The ecosystem grows on demand and is organized around two published package categories:

- **Export** — turn application data into downloadable documents. Excel export ([`@marcusok/excel-exporter`](/packages/excel-exporter/)) is available today; other document formats (e.g. PDF) may follow.
- **Document preview** — preview documents in the browser. A read-only xlsx preview ([`@marcusok/excel-preview`](/packages/excel-preview/)) is available today; other document formats may follow.

Common capability layers (engine, overlay UI) stay private inside the repo and ship bundled into the packages above — keeping the npm surface at exactly what applications install.

Both published packages are documented on this site, so the [package count on the home page](/guide/#current-packages) matches the table above. For which package a given task needs — and what each bundles under the hood — see [Package Relationships & Selection](./03-package-relationships).

When a new package ships, register it in `apps/docs/.vitepress/registry.ts` with its category (`export` / `preview` / `shared`) and add its docs; it automatically appears in the navigation, sidebar and home cards.
