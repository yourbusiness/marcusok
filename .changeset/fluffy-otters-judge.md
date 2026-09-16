---
"@marcusok/excel-exporter": patch
---

Fix cross-path input handling and degradation-chain edge cases found in a full audit:

- Node/SSR stream-route build failures now fail directly instead of re-running the identical fast-stream fallback on the same input (parity with the worker-chain guard) — no more doubled build time and duplicated warnings on a doomed retry.
- Pre-flight validation now covers numeric fields: `width` (finite, non-negative; `0` stays legal and hides the column), `freezeRows` (non-negative integer), `format.decimals` (integer 0–100), and `padding.length` (integer 0–10000). Previously a NaN/negative `width` or fractional `freezeRows` failed the Workbook path with an opaque engine error and degraded the whole export to the style-less stream, while the same input was silently ignored (or threw a `toFixed` RangeError) on the stream path — the same config could succeed below 50k rows and fail above them.
- `exportTable` rejects circular `children` with a clear error instead of a stack overflow (the adapter's recursive mapping ran before the generic cycle check).
- `configureWasm`/`updateOptions` compare `wasmUrl` by string, so re-passing an equal `URL` object no longer resets an already-configured loader; the idempotence warning now also covers the "initial fetch still in flight" case (an in-flight `initWasm` cannot be redirected — documented in the caveat).
- `WasmLoader` state writes are centralized with the promise-identity check: a superseded load (URL changed during a retry backoff) can no longer overwrite the successor's ready state, so `getWasmLoader().isReady` no longer flips back to false.
- Non-`Error` rejections keep their original message in the WASM load error and in worker failure responses (previously `undefined` / "worker unknown error").
- Build: the wasm copy step moved from the `build` script into tsup's `onSuccess` hook, so `tsup --watch` (`pnpm dev`) rebuilds restore `dist/modern-xlsx.wasm` after the clean step wipes it — dev sessions no longer break the Node auto-init path and integration tests.
- Types: corrected `onPhase`/`ExportResult.duration` docs (worker-route duration includes main-thread serialization and the worker round-trip, not in-worker time only) and `autoFilter` docs (the filter range spans the last header row plus all data rows).
- Tests: new direct coverage for the worker message protocol (success path, init idempotency via the string-normalized URL key, byte transfer, error stringification), which previously had only the timeout branch tested.
