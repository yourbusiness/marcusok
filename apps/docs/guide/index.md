# Ecosystem

`marcusok` is a pnpm + Turborepo frontend monorepo providing shared capability packages for multiple admin applications. Each published package is versioned and released independently via Changesets, wired together with `workspace:*` during development.

## Current packages

| Package                                                 | Category | Status | Description                                                                                                   |
| ------------------------------------------------------- | -------- | ------ | ------------------------------------------------------------------------------------------------------------- |
| [`@marcusok/excel-exporter`](/packages/excel-exporter/) | Export   | stable | Excel export engine: modern-xlsx + Fast stream, full styling, Worker threading, fast writes, stream fallback  |
| [`@marcusok/excel-preview`](/packages/excel-preview/)   | Preview  | stable | Read-only xlsx preview: worker parsing + virtual scrolling, styles/merges/freeze/number formats restored      |
| [`@marcusok/word-preview`](/packages/word-preview/)     | Preview  | beta   | Read-only docx preview: visual-faithful rendering (docx-preview engine), format sniffing, hardened defaults   |
| [`@marcusok/pdf-preview`](/packages/pdf-preview/)       | Preview  | beta   | Read-only PDF preview: lazy continuous scroll (pdf.js), selectable text layer, worker/cmaps shipped in `dist` |

Two more capability layers exist in the repo as **private internal packages** — the modern-xlsx engine layer (`xlsx-core`: WASM loading, asset distribution, stable re-export surface) and the progress overlay UI (`progress-overlay`, behind the exporter's `overlay` option). They are bundled into the two Excel business packages' `dist` at build time (the engine layer into both Excel packages, the overlay into the exporter only) and are intentionally not on npm: installing a business package brings everything it needs.

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

- **Export** — turn application data into downloadable documents. Excel export ([`@marcusok/excel-exporter`](/packages/excel-exporter/)) is available today; other document formats may follow.
- **Document preview** — preview documents in the browser. Read-only xlsx ([`@marcusok/excel-preview`](/packages/excel-preview/)), docx ([`@marcusok/word-preview`](/packages/word-preview/)) and PDF ([`@marcusok/pdf-preview`](/packages/pdf-preview/)) previews are available today; other document formats may follow.

Common capability layers (engine, overlay UI) stay private inside the repo and ship bundled into the packages above — keeping the npm surface at exactly what applications install.

All published packages are documented on this site, so the [package count on the home page](/guide/#current-packages) matches the table above. For which package a given task needs — and what each bundles under the hood — see [Package Relationships & Selection](./03-package-relationships).

When a new package ships, three steps put it on this site: add it to `apps/docs/package.json` dependencies (the registry imports each package's `package.json` for its version — the build fails without the dependency), create its pages under `apps/docs/packages/<dir>/`, and register it in `apps/docs/.vitepress/registry.ts` with its category (`export` / `preview` / `shared`); it then automatically appears in the navigation, sidebar and home cards.
