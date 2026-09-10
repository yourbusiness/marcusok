# Node / SSR Usage

In Node servers (including SSR) you don't need browser assets and there is **no initialization boilerplate**: the engine locates this package's `dist/modern-xlsx.wasm` on disk (pnpm-symlink-safe) and initializes it synchronously on first use. The explicit `initWasmSync` bootstrap from earlier versions is no longer required — keep it only if you want the one-off read+compile at startup instead of the first request.

## Environment differences

| Dimension         | Browser              | Node / SSR                                 |
| ----------------- | -------------------- | ------------------------------------------ |
| Worker path       | available            | no Web Worker; falls back to main/stream   |
| Auto download     | triggers download    | `triggerDownload` is a no-op               |
| `download` option | defaults to true     | set `false` explicitly and handle the Blob |
| Large data        | worker + Fast stream | main → stream at ≥ 50k rows (main thread)  |

## Export and write to disk

```ts
import { exportExcel } from "@marcusok/excel-exporter";
import { writeFile } from "node:fs/promises";

const result = await exportExcel({
  filename: "server-report",
  download: false, // never trigger a browser download server-side
  sheets: [{ name: "Sheet1", columns: [...], data: [...] }],
});

if (result.success && result.blob) {
  const buffer = Buffer.from(await result.blob.arrayBuffer());
  await writeFile("./server-report.xlsx", buffer);
}
```

## With a framework (Next.js Route Handler)

```ts
// app/api/export/route.ts
import { exportExcel } from "@marcusok/excel-exporter";

export async function GET() {
  const result = await exportExcel({
    filename: "report",
    download: false,
    sheets: [/* ... */],
  });
  if (!result.success || !result.blob) {
    return Response.json({ error: result.error?.message }, { status: 500 });
  }
  return new Response(result.blob, {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": 'attachment; filename="report.xlsx"',
    },
  });
}
```

## Optional: explicit init timing

To move the one-off synchronous WASM read+compile from the first request to process startup, await the loader before serving traffic:

```ts
import { getWasmLoader } from "@marcusok/excel-exporter";

await getWasmLoader().ensureLoaded(); // reads + compiles the shipped wasm once
```

> Do not use `initWasmSync` from a separately installed `modern-xlsx` for this: the engine is bundled into `@marcusok/excel-exporter`, so an external copy initializes a different module instance and does not pre-warm the bundled one.

> Bundler caveat: if your server build bundles this package and the WASM asset is not emitted alongside the bundle, the automatic disk lookup fails and WASM-dependent routes degrade to the style-less stream (headers/merges preserved). Either keep the package external (the default for Node server builds), pass `configureWasm({ wasmUrl })` with an HTTP URL, or copy the asset where the bundle can read it.

## Performance tip

Large server-side exports (≥ 50k rows) automatically take the stream path; without a Worker, Fast stream occupies the current thread for ~0.8s. Run it in an async task or queue so request threads stay responsive.
