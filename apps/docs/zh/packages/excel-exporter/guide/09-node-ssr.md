# Node / SSR 使用

Node 服务端（含 SSR）无需部署浏览器静态资源，且**无需初始化样板**：未配置 `wasmUrl` 时，首次导出会自动定位安装目录旁的 `dist/modern-xlsx.wasm`（对 pnpm 符号链接安全）并同步初始化（`initWasmSync`），随后正常走带样式的 modern-xlsx 引擎。

> 自动初始化包含一次性的同步文件读取与 WASM 编译，实测约 20ms（读取 1.9MB 二进制约 4ms，编译约 15ms，Node 22 / 本仓库开发机），发生在首次导出时。若仍希望把它提前到进程启动期，用下方显式初始化。

## 环境差异

| 维度            | 浏览器               | Node / SSR                           |
| --------------- | -------------------- | ------------------------------------ |
| Worker 路径     | 可用                 | 无 Web Worker，自动回退 main/stream  |
| 自动下载        | 触发浏览器下载       | `triggerDownload` 为 no-op           |
| `download` 参数 | 默认 true            | 建议显式 `false`，自行处理 Blob      |
| 大数据量        | worker + Fast stream | main → ≥ 5 万行 stream（主线程执行） |
| WASM 初始化     | 自动定位（默认）     | 自动定位并初始化                     |

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

希望把一次性的同步读取+编译从首个请求提前到进程启动期，在服务流量前 await 一次加载器：

```ts
import { getWasmLoader } from "@marcusok/excel-exporter";

await getWasmLoader().ensureLoaded(); // 一次性读取并编译随包发布的 wasm
```

> 不要用单独安装的 `modern-xlsx` 的 `initWasmSync` 来预热：引擎已打包进 `@marcusok/excel-exporter`，外部副本是另一个模块实例，预热不到打包内的这一份。

也可以通过 `configureWasm({ wasmUrl })` 指定一个可 fetch 的 HTTP 地址（走网络加载而非磁盘读取）。

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

> 注意：自动初始化依赖运行时能从磁盘定位安装目录里的 wasm。若你的打包/部署形态不满足这一点（例如依赖被内联进产物且资产未随行输出），自动定位失败时不会报错，而是走降级链（无样式流式兜底，console 有 `[excel-exporter]` 前缀警告）；此时保持本包 external（Node 服务端构建的默认行为）、`configureWasm({ wasmUrl })` 指向 HTTP 地址，或把资产拷贝到产物可读的位置即可。
