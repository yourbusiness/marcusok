# @marcusok/progress-overlay

框架无关的全屏进度遮罩：任务时长未知时显示旋转圆环，收到真实进度后切换为细百分比条。毛玻璃面板、明暗双主题、零运行时依赖——它是 [@marcusok/excel-exporter](/zh/packages/excel-exporter/) `overlay` 选项背后的共享 UI 层，也可独立用于任何耗时任务（报表生成、批量上传、数据导入……）。

## 安装

```bash
pnpm add @marcusok/progress-overlay
```

零运行时依赖——纯 DOM 加一次注入的样式表。

## 快速示例

```ts
import { showProgressOverlay } from "@marcusok/progress-overlay";

const overlay = showProgressOverlay({
  text: {
    title: "正在生成报表",
    phases: { building: "构建工作簿…", downloading: "下载中…" },
  },
});

try {
  await runTask({
    onProgress: (p) => overlay.setProgress(p), // 0..1
    onStage: (key) => overlay.setPhase(key), // text.phases 里的 key
  });
} finally {
  overlay.close(); // 幂等；唯一正确的关闭时机
}
```

## 能力

| 能力         | 说明                                                                                                                                                                                       |
| ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 两态自适应   | 时长未知显示圆环 + 提示；`(0, 1)` 区间的进度值一到即切百分比条。不伪造：仅有一个收尾的 `1` 绝不会凭空补出走完的进度条。                                                                    |
| 中性玻璃观感 | zinc 灰阶不与宿主品牌色冲突；模糊蒙层、带内高光的玻璃面板、`prefers-color-scheme` 主题（`auto` / `light` / `dark`）。                                                                      |
| 延迟门控     | 默认 `delayMs: 200`——比它更快的任务完全不闪遮罩；最短可见时长避免一闪而过。                                                                                                                |
| 并发         | 并发任务共享一个引用计数的 DOM 节点（最后更新者为准）；文案始终属于最近介入且仍活动的调用方。                                                                                              |
| 健壮性       | 句柄的每个方法都保证不抛错（遮罩绝不允许弄挂它装饰的任务）；Node/SSR 返回空实现句柄——同一处调用代码，无需分支。                                                                            |
| 无障碍       | 进度条带 `role="progressbar"`（仅确定态附带 `aria-valuenow`）；文案节点是 `aria-live="polite"` 状态区域且百分比独立在外（不会每次进度跳动都重播）；`prefers-reduced-motion` 停用全部动效。 |

## 文案从哪来

本包不含业务词汇：默认文案是通用的（`请稍候` / `正在处理…`），阶段文案是调用方定义的 `key → 文案` 表（`text.phases`），未覆盖的 key 原样显示（拼写错误立刻可见）。`@marcusok/excel-exporter` 把它的 `ExportPhase` 事件映射到这套协议并提供"正在导出 Excel"文案——见[导出包的遮罩指南](/zh/packages/excel-exporter/guide/11-overlay)。

完整选项表、阻塞线程的取舍（`delayMs: 0` + `nextPaint()`）与并发语义见[使用指南](./guide/01-usage)。
