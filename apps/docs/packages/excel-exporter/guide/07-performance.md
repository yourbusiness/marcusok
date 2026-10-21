# Performance

Numbers below come from local measurements (real Chrome, 6 mixed-type columns). The Node **in-process** regression baseline (SLA 200/500/1000 ms) lives in `src/__tests__/performance.test.ts`; the toBuffer cliff is verified out-of-band via independent processes (see the package README) — measured in Node, the same 100k Workbook cliff lands around ~21s cold, versus the 17.5s Chrome figure below: one cliff, two environments.

## Benchmarks

| Rows    | Workbook (main) | Fast stream | What `auto` picks    |
| ------- | --------------- | ----------- | -------------------- |
| 10,000  | ~120ms          | —           | main + Workbook      |
| 50,000  | —               | ~400ms      | Worker + Fast stream |
| 100,000 | 17.5s (legacy)  | ~780ms      | Worker + Fast stream |

<ClientOnly>
  <BenchmarkChart dir="excel-exporter" />
</ClientOnly>

## Takeaways

1. `Workbook.toBuffer()` shows a superlinear cliff beyond ~55k rows (17.5s at 100k), while Fast stream stays at ~0.8s — hence `STREAM_THRESHOLD = 50_000`;
2. In the browser, ≥ 20,000 rows run in a Worker; the main thread only does one structured clone (~94ms at 100k rows);
3. The stream path's cost is missing styles/layout, so small files keep the fully-styled Workbook path.

## Optimization tips

- Stay on `mode: "auto"`; never force `main` for large files;
- Keep styled exports under 50k rows **per call** — `auto` routes on the total across all sheets of one call, so splitting into more sheets of the same call does not help; for larger data accept degraded styling or split into multiple `exportExcel` calls (one file each);
- `decimals`/`thousands` are display-only numFmt settings on both paths now; no need to pre-round values for cross-threshold consistency;
- Run large server-side exports in background tasks.

> Runtime note: the local baselines and `.nvmrc` use Node 24.21.0. The bundled engine's WASM `readBuffer` is ~10x slower on Node 22.x (and Node 24.14) for numeric-heavy preview sheets; the exporter's Workbook/stream benchmarks are far less sensitive, but performance comparisons should use the pinned runtime. CI keeps a functional Node 22 job with performance tests skipped.
