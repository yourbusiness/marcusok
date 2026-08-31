# CLAUDE.md

## 交流约定

- 与用户的所有交流、回复必须使用中文（代码、提交信息、用户可见文档仍按下方文档约定执行）。

## 仓库速览

pnpm 9 + Turborepo monorepo，Node ≥ 22（`.npmrc` 设 `engine-strict=false`，勿改，原因见 README）。
工具链、目录布局、发版细节见 README.md 与 docs/，此处不重复。

## 常用命令

- `pnpm build / test / lint / typecheck`（turbo 编排）
- 单包验证：`pnpm --filter @marcusok/excel-exporter test`
- `pnpm dev:play` / `pnpm dev:docs`（本地调试）

## Git 工作流（硬性约束）

- **改动代码前**：必须先 `git pull` 同步远端代码。
- **改动完成后**：未得到用户明确确认前，禁止执行 `git add` / `git commit` / `git push`。
- **用户确认可提交后**：最先判断本次改动是否需要发版（是否影响 `packages/` 下发布包）；若需要，先 `pnpm changeset` 生成变更集（版本号与 CHANGELOG 由 Changesets 管理，禁止手动修改），再执行后续 git 操作。
- 直接在 `main` 分支开发提交，不建功能分支。
- 提交信息必须 Conventional Commits（commitlint + husky 强制）。

## 技术红线

- 全仓 ESM-only（`type: "module"`），不写 CJS。
- 库打包用 tsup，测试用 Vitest，不引入其他构建器/测试框架。
- 修改发布包代码必须配套 changeset（见 Git 工作流）。

## 文档约定

- 用户可见文档（README、`apps/docs` 文档站）用英文；`docs/` 内部设计文档用中文。
- 修改 excel-exporter 功能时，同步检查：包 README、`apps/docs` 文档站、`packages/play` 演示。

## 仓库事实

- GitHub owner `yourbusiness` 是真实账号，非占位符。
- 文档站：https://yourbusiness.github.io/marcusok/
