# @marcusok/xlsx-core

Shared engine layer for the [@marcusok](https://github.com/yourbusiness/marcusok) spreadsheet packages
([excel-exporter](https://www.npmjs.com/package/@marcusok/excel-exporter),
excel-preview). It is the repo's single integration point around
[modern-xlsx](https://www.npmjs.com/package/modern-xlsx) (Rust + WASM):

- **WASM loading** — `configureWasm()` / `getWasmLoader()` with a browser
  fetch / Node sync-init dual path, per-attempt timeout and retries. Zero
  configuration in both environments.
- **Asset distribution** — the WASM binary ships in `dist/` and is
  re-published under this package's exports map
  (`@marcusok/xlsx-core/dist/modern-xlsx.wasm?url` works in Vite/webpack 5).
- **A stable re-export surface** — the modern-xlsx APIs the @marcusok
  packages share (`readBuffer`, `formatCellRich`, date-serial utilities,
  reference utilities, typed errors). The engine's runtime is bundled in;
  the declared `modern-xlsx` dependency exists for d.ts type resolution only.
- **tsup esbuild plugins** — `@marcusok/xlsx-core/tsup` exposes the two
  plugins (wasm-glue URL rewrite, Node-only `node:fs/promises` drop) other
  packages reuse when bundling self-contained workers.

The engine's runtime is bundled **into** this package: nothing imports
modern-xlsx at runtime, so its `engines.node>=24` range is inert at runtime.
The package does declare one dependency — `modern-xlsx`, pinned at the exact
bundled version — solely so TypeScript consumers can resolve the re-exported
types in the published d.ts (tsup leaves type re-exports as external
`from "modern-xlsx"` imports — modern-xlsx's exports-map-only package shape
has no top-level `types` entry, which tsup's dts resolver cannot resolve,
verified 2026-09: neither devDependency placement nor import+re-export
syntax changes this; without the declaration they fail
with TS2307 in projects that never installed it). Caveat: the declaration
does make the pinned version visible to installers — with engine-strict
enabled on Node < 24 the install is rejected (npm/pnpm default: warning
only). Packages consume it
with `external` in their main builds, so a page using several @marcusok
packages loads **one** engine instance and **one** WASM binary on the main
thread (their self-contained workers bundle their own copy by necessity —
browser module workers cannot resolve bare specifiers).

## Install

```bash
pnpm add @marcusok/xlsx-core
# or npm i @marcusok/xlsx-core / yarn add @marcusok/xlsx-core
```

## Usage

Usually consumed indirectly through `@marcusok/excel-exporter` (or
excel-preview). Direct use:

```ts
import { configureWasm, readBuffer } from "@marcusok/xlsx-core";

// Optional: point at a self-hosted wasm copy. Defaults to the shipped
// binary next to the entry (bundler-emitted or node_modules-resolved).
configureWasm({ wasmUrl: "/assets/modern-xlsx.wasm" });

const workbook = await readBuffer(bytes);
```

Node needs no configuration at all — the loader locates and synchronously
initializes the shipped binary on first use.

## Why a separate core?

Before this package, each business package bundled its own copy of
modern-xlsx. Two packages on one page then shipped two engine instances and
could load the 2 MB WASM twice. With the engine behind this shared core,
the duplication on the main thread is gone while each package's public API
stays unchanged.

## License

MIT
