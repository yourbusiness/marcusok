# 生态介绍

`marcusok` 是一个 pnpm + Turborepo 组织的前端基础 monorepo，为多个后台应用提供统一的公共能力包。每个包独立版本、独立发布（Changesets），通过 `workspace:*` 在仓库内互引。

## 当前包

| 包                                                         | 分类 | 状态   | 说明                                                                                         |
| ---------------------------------------------------------- | ---- | ------ | -------------------------------------------------------------------------------------------- |
| [`@marcusok/excel-exporter`](/zh/packages/excel-exporter/) | 导出 | stable | Excel 导出核心库：modern-xlsx + Fast stream、完整样式、Worker 多线程、快速写入、流式降级兜底 |
| [`@marcusok/excel-preview`](/zh/packages/excel-preview/)   | 预览 | stable | xlsx 只读预览：Worker 解析 + 虚拟滚动，样式/合并/冻结/数字格式还原                           |

仓库里还有两个**私有内部包**作为能力层——modern-xlsx 引擎层（`xlsx-core`：WASM 加载、资产分发、稳定再导出）与进度遮罩 UI（`progress-overlay`，导出包 `overlay` 选项的底层）。它们在构建期整体打进发布包的 `dist`，有意不发布到 npm：安装业务包即得全部能力。

## 工程约定

- 包管理器：pnpm workspace（`pnpm >= 9`）
- 构建编排：Turborepo（`^build` 自动处理包依赖顺序）
- 包构建：tsup（TS → ESM + DTS，ESM-only）
- 语言：TypeScript 5.x（`moduleResolution: bundler`）
- 测试：Vitest；代码规范：ESLint 9 + Prettier
- 版本/发布：Changesets（多包独立发版、自动 changelog、支持 prerelease）
- CI/CD：GitHub Actions（`ci.yml` 校验、`release.yml` 自动发布 npm、`deploy.yml` 部署本文档站到 GitHub Pages）

## 路线图

生态将按需扩展，按两大发布类组织：

- **导出** —— 把业务数据导出为可下载的文档。Excel 导出（[`@marcusok/excel-exporter`](/zh/packages/excel-exporter/)）已交付，后续可能补充其他文档类型（如 PDF）。
- **文档预览** —— 在浏览器内预览文档。只读 xlsx 预览（[`@marcusok/excel-preview`](/zh/packages/excel-preview/)）已交付，其他文档格式以后按需跟进。

公共能力层（引擎、遮罩 UI）保持仓库私有、随上述包内置分发——npm 发布面收敛为应用真正要安装的东西。

两个已发布包都已收录到本站，因此[首页统计的包数量](/zh/guide/#当前包)与上表一致。至于某个任务该用哪个包、它们各自内置了什么，见[包关系与选型](/zh/guide/03-package-relationships)。

新包发布后，只需在文档站 `apps/docs/.vitepress/registry.ts` 登记一条记录（含所属分类：`export` / `preview` / `shared`）并补充对应文档，即可自动出现在本站在线文档中。
