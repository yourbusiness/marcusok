# Fallback (style-less fast stream)

When the WASM path is unavailable, the library automatically degrades to the pure-JS fast stream, so most failures still produce a file (one exception is listed under "When it triggers"). The fallback needs no WASM, no Worker and no network access — it is the same stream writer used for ≥ 50,000-row exports, run on the main thread without styles.

## When it triggers

- The environment does not support `WebAssembly` (this only affects the main and Worker + Workbook paths; the ≥ 50,000-row stream path does not use WASM and is unaffected);
- `modern-xlsx.wasm` fails to load (after `maxRetries` attempts, default 3);
- The Worker path fails (e.g. the worker asset 404s) **and** the automatic main-thread retry also fails — on the Workbook route the retry runs modern-xlsx on the main thread first (styles preserved), so the stream is the last resort, not the immediate next step. On the ≥ 50,000-row stream route the retry is the fast stream itself (identical code, identical input), so a failure there is terminal — no third attempt;
- The build throws on a main-thread Workbook route (e.g. an internal modern-xlsx build error). The same error on a Node stream route (no `window`, explicit `mode: "stream"` or auto ≥ 50,000 rows) is terminal — the first attempt already ran the fast stream on the identical input, so there is nothing to retry or fall back to. Note: structural input errors (invalid/duplicate sheet names, out-of-bounds merges, non-numeric or negative `width` / fractional `freezeRows`, out-of-range `format.decimals`, …) never enter this chain — pre-flight validation rejects them before any route runs, so the export resolves with `success: false` without a fallback attempt.

## Behavioral differences

| Dimension                  | modern-xlsx path | stream fallback                                                                           |
| -------------------------- | ---------------- | ----------------------------------------------------------------------------------------- |
| `ExportResult.engine`      | `"modern-xlsx"`  | `"modern-xlsx"` with `mode: "stream"` and `result.error` set (success stays `true`)       |
| Cell styles                | full             | stripped (the fast stream emits minimal OOXML without a style part)                       |
| Multi-row headers / merges | supported        | supported (merges are structure, not styles)                                              |
| Width / freeze / filter    | supported        | not supported                                                                             |
| FormatSpec                 | supported        | supported (enum/padding/number/date semantics kept; dates become readable strings)        |
| Number formats             | `numFormat`      | `decimals` baked into stored value                                                        |
| Warning                    | —                | console prints `[excel-exporter] Falling back to the style-less fast stream. Reason: ...` |

## Detecting the fallback

```ts
const result = await exportExcel(options);
if (result.success && result.error) {
  // Degraded export: styles stripped, reason in result.error.message
}
```

The fallback is an insurance path for failures, not a regular one. When it fires, first check that the wasm asset is not 404 (open the Network tab) and — if you use `configureWasm` overrides — that the URLs were set before the first export. A subtler variant: the wasm request shows **200/304 yet the response is HTML** (`content-type: text/html`; the underlying compile error mentions `expected magic word 00 61 73 6d, found 3c 21 64 6f`). That is Vite's dev-server dependency pre-bundling answering from its HTML fallback — the fix is an `optimizeDeps.exclude` entry, see [Installation → Vite dev server](/packages/excel-exporter/guide/02-installation#vite-dev-server-pre-bundling-caveat). Also note: after the first failed attempt, later exports report `WASM load previously failed` instead of the original error — the real reason only appears in the console warning of the **first** export (or after a page reload). The one exception is a late success on a slow network: if the underlying load eventually settles after all retries timed out, the loader heals itself back to ready and later exports keep the fully-styled paths (no `configureWasm` needed).
