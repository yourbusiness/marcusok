# 在线演示

本页所有演示都在浏览器内直接运行发布态的库包——解析、构建与进度全在本地完成，不上传任何数据。数据由文档站的确定性 mock 生成器产生（mulberry32 种子 PRNG），同一配置多次生成结果一致，方便复现。

## Excel 导出 —— `@marcusok/excel-exporter`

选择数据集、数据量与导出模式，点击导出即可得到真实 `.xlsx` 文件，同时展示进度、各阶段耗时与最终引擎信息。

<ClientOnly>
  <PackageDemo dir="excel-exporter" />
</ClientOnly>

- **auto（推荐）**：按数据量自动选择 main / worker / stream 最优路径；
- **main**：主线程同步构建，10 万行时可以看到明显的性能断崖；
- **worker**：Web Worker 多线程，主线程只做一次结构化克隆；
- **stream**：Fast stream，10 万行约 0.8s，但不支持样式与布局特性。

完整用法见 [excel-exporter 文档](/zh/packages/excel-exporter/)。

## xlsx 预览 —— `@marcusok/excel-preview`

把 `.xlsx` / `.xlsm` / `.csv` 文件拖入面板：解析在 Web Worker 内完成，表格走 DOM 虚拟滚动渲染，大文件下面也不会卡顿。文件解析完成后会给出各阶段耗时（解析 / 渲染 / 合计）；加密文件会明确报 `PASSWORD_PROTECTED`，不会静默失败。

<ClientOnly>
  <PackageDemo dir="excel-preview" />
</ClientOnly>

完整用法见 [excel-preview 文档](/zh/packages/excel-preview/)。

## 进度遮罩 —— `@marcusok/progress-overlay`

导出过程中看到的那个遮罩，这里不经任何业务包直接驱动，任务与 Excel 无关。模拟任务前 1.5 秒不上报进度——那就是不确定态的旋转圆环；随后流式推进并切换为百分比条。两态渲染的是同一个共享组件、同一个毛玻璃面板。

<ClientOnly>
  <PackageDemo dir="progress-overlay" />
</ClientOnly>

完整用法见 [progress-overlay 文档](/zh/packages/progress-overlay/)。

## 引擎层 —— `@marcusok/xlsx-core`

这个包没有在线演示：`xlsx-core` 是无界面的引擎层——WASM 加载器加上对 modern-xlsx 的稳定再导出面——本身不渲染任何东西。上面两个面板已经在用它：两个业务包共用本包加载的同一个引擎实例与同一份 WASM 二进制。加载器、exports 映射与"何时直接用"见 [xlsx-core 文档](/zh/packages/xlsx-core/)。
