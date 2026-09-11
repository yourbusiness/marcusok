# Fallback (style-less fast stream)

When the WASM path is unavailable, the library automatically degrades to the pure-JS fast stream, so most failures still produce a file (one exception is listed under "When it triggers"). The fallback needs no WASM, no Worker and no network access — it is the same stream writer used for ≥ 50,000-row exports, run on the main thread without styles.

## When it triggers

- The environment does not support `WebAssembly` (this only affects the main and Worker + Workbook paths; the ≥ 50,000-row stream path does not use WASM and is unaffected);
- `modern-xlsx.wasm` fails to load (after `maxRetries` attempts, default 3);
- The Worker path fails (e.g. the worker asset 404s) **and** the automatic main-thread retry also fails — on the Workbook route the retry runs modern-xlsx on the main thread first (styles preserved), so the stream is the last resort, not the immediate next step. On the ≥ 50,000-row stream route the retry is the fast stream itself (identical code, identical input), so a failure there is terminal — no third attempt;
- The build throws (e.g. an internal modern-xlsx build error). Note: an invalid sheet name is not rescuable here — the fallback re-validates the same name and fails, so the export resolves with `success: false`.

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

The fallback is an insurance path for failures, not a regular one. When it fires, first check that the wasm asset is not 404 (open the Network tab) and — if you use `configureWasm` overrides — that the URLs were set before the first export.
