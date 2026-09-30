# @marcusok/xlsx-core（仓库私有包）

> **私有内部包**：不再发布到 npm（3.0 之前曾以 `@marcusok/xlsx-core`
> 名义发布，已停发并废弃）。它是两个业务包
> ([excel-exporter](https://www.npmjs.com/package/@marcusok/excel-exporter) /
> [excel-preview](https://www.npmjs.com/package/@marcusok/excel-preview))
> 的构建期依赖：两个包在 tsup 构建中把本包的 `dist` **整体打进各自的产物**
> （主入口与自包含 worker 都是 `noExternal`）。消费方永远只安装业务包本身，
> 不需要也不应直接依赖本包。

全仓唯一的 [modern-xlsx](https://www.npmjs.com/package/modern-xlsx)
(Rust + WASM) 集成层：

- **WASM 加载** —— `configureWasm()` / `getWasmLoader()`，浏览器 fetch /
  Node 同步初始化双路径、单次超时与重试。两种环境零配置。
- **资产分发** —— WASM 二进制随 `dist/` 分发；业务包各自把它再转发进
  自己的 `dist`（`@marcusok/excel-exporter/dist/modern-xlsx.wasm` 等路径
  即由此而来）。
- **稳定的再导出面** —— 各业务包共用的 modern-xlsx API（`readBuffer`、
  `formatCellRich`、日期序列工具、引用工具、类型化错误）。引擎运行时
  已打包进本包 `dist`。
- **tsup esbuild 插件** —— `@marcusok/xlsx-core/tsup` 暴露两个插件
  （wasm-glue URL 重写、剔除 Node-only 的 `node:fs/promises`），业务包
  的 tsup 配置复用（构建期 devDependency，不进运行时）。

## 类型依赖说明（唯一保留的 dependencies 条目）

本包声明 `modern-xlsx` 依赖（钉在与打包版本完全一致的 1.2.0），**仅为
d.ts 类型解析**：tsup 的 dts pass 会把类型再导出保留为外部的
`from "modern-xlsx"` 导入（modern-xlsx 是纯 exports-map 包、无顶层
`types`，tsup 的 dts resolver 解析不了，2026-09 实测：devDependencies
放置与 import+re-export 写法均不改变该行为），不声明则未安装它的项目
直接 TS2307。同样的声明与理由现随打包一并上移到两个业务包的
`dependencies`。运行时不加载任何 modern-xlsx 代码，其
`engines.node>=24` 对运行时无效（engine-strict 安装器除外——npm/pnpm
默认仅告警）。

## 为什么收敛为私有层

最初每个业务包各自打包 modern-xlsx；后来抽出过共享发布包（同页单引擎）；
最终（exporter 3.0 / preview 2.0 起）为收敛 npm 发版面（4 包 → 2 包、
消除跨包版本联动）转为**私有构建期层**。代价：同页使用两个业务包时
主线程各持一份引擎与 WASM（两份二进制相同，通常被内容 hash 资产与
HTTP 缓存去重）；各包自包含 worker 本来就各带一份，不受影响。

## License

MIT
