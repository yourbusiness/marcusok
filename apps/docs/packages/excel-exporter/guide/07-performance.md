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

Wide sheets are measured separately (Node 24.21): 40k×50 = 2M cells takes ~6.3s / RSS +1.2GB on Workbook vs ~1s / +58MB on Fast stream; 20k×100 = 2M cells takes ~3.2s / +258MB on Workbook; 60k×20 = 1.2M cells is ~1s on stream. Routing therefore gates on estimated cells as well as rows.

## Takeaways

1. `Workbook.toBuffer()` shows a superlinear cliff beyond ~55k rows (17.5s at 100k), while Fast stream stays at ~0.8s — hence the 50,000-row and 1,500,000-cell stream gates;
2. In the browser, ≥ 20,000 rows **or** ≥ 250,000 cells run in a Worker; the main thread only does one structured clone, whose cost grows with cell count (~94ms at 100k×6, ~200ms at 40k×50);
3. The stream path's cost is missing styles/layout, so small files keep the fully-styled Workbook path; `result.degradedFeatures` makes that loss programmatically visible.

## Optimization tips

- Stay on `mode: "auto"`; never force `main` for large files;
- Keep styled exports below **both** gates per call (rows < 50k **and** cells < 1.5M) — `auto` routes on the total across all sheets of one call, so splitting into more sheets of the same call does not help; for larger data accept degraded styling (check `result.degradedFeatures`) or split into multiple `exportExcel` calls (one file each);
- `decimals`/`thousands` are display-only numFmt settings on both paths now; no need to pre-round values for cross-threshold consistency;
- Run large server-side exports in background tasks.

> Runtime note: the local baselines and `.nvmrc` use Node 24.21.0. The bundled engine's WASM `readBuffer` is ~10x slower on Node 22.x (and Node 24.14) for numeric-heavy preview sheets; the exporter's Workbook/stream benchmarks are far less sensitive, but performance comparisons should use the pinned runtime. CI keeps a functional Node 22 job with performance tests skipped.
