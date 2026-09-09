# Node / SSR 使用

Node 服务端（含 SSR）无需部署浏览器静态资源，且**无需初始化样板**：未配置 `wasmUrl` 时，首次导出会自动通过 `createRequire` 从 `node_modules` 定位 `modern-xlsx.wasm` 并同步初始化（`initWasmSync`，对 pnpm 符号链接安全），随后正常走带样式的 modern-xlsx 引擎。

> 自动初始化包含一次性的同步文件读取与 WASM 编译，实测约 20ms（读取 1.9MB 二进制约 4ms，编译约 15ms，Node 22 / 本仓库开发机），发生在首次导出时。若仍希望把它提前到进程启动期，用下方显式初始化。

## 环境差异

| 维度            | 浏览器               | Node / SSR                                |
| --------------- | -------------------- | ----------------------------------------- |
| Worker 路径     | 可用                 | 无 Web Worker，自动回退 main/stream       |
| 自动下载        | 触发浏览器下载       | `triggerDownload` 为 no-op                |
| `download` 参数 | 默认 true            | 建议显式 `false`，自行处理 Blob           |
| 大数据量        | worker + Fast stream | main → ≥ 5 万行 stream（主线程执行）      |
| WASM 初始化     | 配置 `wasmUrl`       | 自动定位并初始化（或显式 `initWasmSync`） |

## 服务端导出并落盘

```ts
import { exportExcel } from "@marcusok/excel-exporter";
import { writeFile } from "node:fs/promises";

const result = await exportExcel({
  filename: "server-report",
  download: false, // 服务端不要触发浏览器下载
  sheets: [{ name: "Sheet1", columns: [...], data: [...] }],
});

if (result.success && result.blob) {
  const buffer = Buffer.from(await result.blob.arrayBuffer());
  await writeFile("./server-report.xlsx", buffer);
}
```

## 显式初始化（可选）

希望控制初始化时机（如进程启动期预初始化），或部署形态导致自动定位不可用时，保留显式写法：

```ts
import { initWasmSync } from "modern-xlsx";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
initWasmSync(
  readFileSync(
    `${require("node:path").dirname(require.resolve("modern-xlsx"))}/modern-xlsx.wasm`,
  ),
);
```

也可以通过 `configureWasm({ wasmUrl })` 指定一个可 fetch 的 HTTP 地址（走网络加载而非磁盘读取）。显式初始化与自动初始化幂等共存：modern-xlsx 内部有共享的"首次成功初始化生效"标志，先到者生效。

## 配合框架（如 Next.js Route Handler）

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

## 性能提示

服务端大文件（≥ 5 万行）会自动走 stream 路径；由于没有 Worker，Fast stream 占用当前线程约 0.8s，适合放在异步任务/队列中，避免阻塞请求线程。

> 注意：自动初始化依赖运行时能从磁盘定位 `node_modules` 里的 modern-xlsx。若你的打包/部署形态不满足这一点（例如依赖被内联进产物），自动定位失败时不会报错，而是走原有降级链（SheetJS 兜底，console 有 `[excel-exporter]` 前缀警告）；此时改用上方显式初始化，或 `configureWasm({ wasmUrl })` 指向 HTTP 地址即可。
