# 在包中接入引擎层

本页面向**构建在** `@marcusok/xlsx-core` 之上的包作者：core 为 `tsup` 提供的两个 esbuild 插件、如何接线、WASM 资产如何定位，以及依赖 core 带来的 `engines` 约束。应用侧的用法请看 [WASM 加载器指南](/zh/packages/xlsx-core/guide/01-loader)。

## 为什么一个包需要构建插件

引擎以 JS 胶水加 `.wasm` 二进制两种形态发布，两者都带着在消费方浏览器构建里会出问题的假设：

- wasm-bindgen 胶水用 `new URL('modern_xlsx_wasm_bg.wasm', import.meta.url)` 定位自己的二进制——而**没有任何** @marcusok 包发布这个文件名（二进制被转发为 `dist/modern-xlsx.wasm`）。运行时这条分支是死代码，因为 loader 总是传显式 URL，但打包器会静态分析每一处字面量 `new URL(..., import.meta.url)`，文件不存在就告警（Vite：_"doesn't exist at build time, it will remain unchanged…"_）。
- modern-xlsx 的 Node 专用文件 API（`toFile`、`readFile`）会动态 import `node:fs/promises`。@marcusok 各包只导出 buffer 系 API（`toBuffer` / `readBuffer`），因此这些 import 是死分支——但它们会以字面量动态 import 的形式留在产物里，消费方浏览器构建为此告警（Vite 5 / VitePress：_"Module fs/promises has been externalized for browser compatibility"_）。

`@marcusok/xlsx-core/tsup` 为每个问题导出一个插件，二者都在**你的包的构建期**生效，坏代码因此根本到不了消费方的打包器。

| 插件                 | 改写什么                                                                                                                                       | 为何安全                                                                                                                                                                                                                            |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `rewriteWasmBgUrl`   | 胶水文件里的 `new URL('modern_xlsx_wasm_bg.wasm', import.meta.url)` → `new URL(["..", "dist", "modern-xlsx.wasm"].join("/"), import.meta.url)` | 替换结果刻意**不是**字符串字面量：esbuild 与下游打包器都不匹配计算式说明符，不会被资产分析、也不会重复发射。即便该分支真的执行也能解析成功——从胶水文件与从你打包产物的 `dist` 出发，`../dist/modern-xlsx.wasm` 都是已发布的一份拷贝 |
| `dropNodeFsPromises` | modern-xlsx 自己 dist chunk 里的 `await import('node:fs/promises')` → `await Promise.reject(new Error(...))`                                   | 引擎的 Node 专用 API 不在 @marcusok 包的导出面内；一个死分支以明确的错误失败，好过留下"被外部化的模块桩"这种谜题                                                                                                                    |

第二个插件的**作用域**很重要：过滤规则只匹配 modern-xlsx 的 dist chunk。你自己那处真正会被执行的 Node 内置模块 `await import(...)` 必须保留——core 的 Node 自动初始化正是这么用的，用计算式说明符让浏览器打包器永远看不到静态 `node:fs` 导入。

## 在 tsup.config.ts 中接入

两个消费包的写法完全一致——把共享选项对象展开进每份配置，worker 入口因此也带上插件：

```ts
import { defineConfig, type Options } from "tsup";
import { rewriteWasmBgUrl, dropNodeFsPromises } from "@marcusok/xlsx-core/tsup";

const shared: Partial<Options> = {
  esbuildPlugins: [rewriteWasmBgUrl, dropNodeFsPromises],
};

export default defineConfig([
  {
    ...shared,
    entry: { index: "src/index.ts" },
    format: ["esm"],
    dts: { resolve: true },
    // 保持 core external：同页使用多个 @marcusok 包时，主线程只加载一份引擎与一份 WASM
    external: ["@marcusok/xlsx-core"],
    platform: "browser",
    // 每次构建、以及 watch 的每次重建后都回补二进制（见下文）
    onSuccess: "node scripts/copy-wasm.mjs",
  },
  {
    ...shared,
    entry: { "my.worker": "src/workers/my.worker.ts" },
    format: ["esm"],
    dts: false,
    splitting: false, // 单文件自包含——见下文
    platform: "browser",
    // worker 把 core 打包进来：浏览器 module worker 无法解析裸导入，脚本必须自包含
    noExternal: ["@marcusok/xlsx-core"],
    clean: false,
  },
]);
```

`@marcusok/xlsx-core/tsup` 是**仅 Node** 的入口（构建期用 `node:fs` 读取胶水文件），因此它只应出现在 tsup 配置里，由你包的 devDependency 解析。由于 core 把引擎一并打包，你**不需要**把 `modern-xlsx` 列为依赖，也不需要 `noExternal: ["modern-xlsx"]`——`noExternal: ["@marcusok/xlsx-core"]` 会把整份引擎带进来。

**主入口 external、worker 打包进来。** 这个分工正是 core 存在的意义：主线程在整页 @marcusok 包之间共享一份引擎实例（与一份 WASM 二进制），而 worker 脚本必须是单文件自包含。构建后用下面这条命令验证：

```bash
grep -c "^import" dist/my.worker.js   # 必须为 0
```

分块产出但兄弟导入未被跟踪的 worker 会在生产构建里 404，这也是 worker 入口关闭 `splitting` 的原因。

## WASM 资产的定位与发布

loader 的默认 URL 是 `new URL("./modern-xlsx.wasm", import.meta.url)`——构建后二进制必须位于你包入口旁边。用构建后拷贝步骤转发它，理由与 core 自己的一样：

```js
// scripts/copy-wasm.mjs
import { copyFileSync, mkdirSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const distDir = resolve(dirname(fileURLToPath(import.meta.url)), "..", "dist");
const src = resolve(
  dirname(require.resolve("modern-xlsx")),
  "modern-xlsx.wasm",
);

if (!statSync(src, { throwIfNoEntry: false })) {
  throw new Error(
    `modern-xlsx.wasm not found at ${src}. Run pnpm install first.`,
  );
}

mkdirSync(distDir, { recursive: true });
copyFileSync(src, resolve(distDir, "modern-xlsx.wasm"));
```

`require.resolve("modern-xlsx")` 解析到包入口，其所在目录即二进制位置——与 loader 的 Node 自动初始化用的解析方式相同。注意这暗示了什么：两个消费包都把自己声明 `modern-xlsx` 为 **devDependency**，纯粹是为了让这个脚本能找到源二进制（运行时它们解析的始终是 `@marcusok/xlsx-core`，从不解析 `modern-xlsx`）。在 pnpm 的严格 `node_modules` 布局下，core 的依赖从你的包里解析不到，所以脚本借不到 core 那一份。

另外两条本仓库踩过坑的经验：

- 挂在 tsup 的 `onSuccess` 上，而不只是挂在 `build` npm 脚本上。主配置用了 `clean: true`，`--watch` 启动时会清空 `dist/`（连带上次拷贝的 wasm）且此后不再回补；`onSuccess` 在 `build` 与 `dev` 两条链路的每次重建后都会执行。
- 把它发布在你包的 `exports` 映射下（例如 `"./dist/modern-xlsx.wasm": "./dist/modern-xlsx.wasm"` 这样的文件条目）。这才让 `你的包/dist/modern-xlsx.wasm?url` 在 Vite / webpack 5 中可解析——引擎自己的 exports 映射刻意不含 wasm 子路径，所以深层导入 `modern-xlsx` 走不通。

`new URL(<文件>, import.meta.url)` 随后就是唯一需要的定位方式：打包器改写该表达式并把文件发射为 hash 资产，Node 经 `node_modules` 解析到它。要在 Node 里从磁盘读取（loader 自动初始化做的事），还要求包的 `dist/` 确实在磁盘上——把依赖内联进产物却不发射资产的构建形态，必须保持包 external、用 `configureWasm({ wasmUrl })` 指向 HTTP 地址，或把该文件拷到产物旁。

## engines 与钉住的依赖

- core 声明 `engines.node >= 22`。它打包进来的引擎声明的是 `>= 24`，但该范围在运行时是惰性的——没有任何代码 import `modern-xlsx`——这正是 core 重新声明一个消费方能满足的范围的原因。
- 尽管如此，`modern-xlsx` 仍以精确钉住的版本留在 core 的依赖里，**只为解析类型**：tsup 的 dts 流程无法内联再导出的类型（modern-xlsx 是纯 exports 映射包、没有顶层 `types` 条目，tsup 的 dts resolver 解析不了它），因此发布的 `.d.ts` 会保留 `from "modern-xlsx"` 形式的导入，TypeScript 需要该声明才能解析。不要把它当"没用的依赖"清理掉。
- 由此给安装方的副作用：在 Node < 24 上开启 `engine-strict=true` 时安装会被拒绝（npm / pnpm 默认只是警告）。完整缘由见[包 README](https://www.npmjs.com/package/@marcusok/xlsx-core)。

## 处理引擎错误

如果你的包要把引擎失败面暴露给调用方，请按 `ModernXlsxError` 的 `code` 分流而不是按 message 文本——消息随输入形态变化，码才是稳定契约。把你关心的码映射到自己的错误词汇上，`@marcusok/excel-preview` 的[错误归一模块](/zh/packages/excel-preview/api/01-create-preview#previewerror-错误码)就是这么做的。以值形式导出的四个常量（[类型化错误](/zh/packages/xlsx-core/guide/02-engine-surface#类型化错误)）覆盖了最要紧的几类；其余只能按字面字符串匹配。
