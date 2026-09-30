# API：showProgressOverlay 与外观定制

本页是这两个导出与视觉面的参考。关于它们背后的概念——遮罩何时出现、句柄在并发下的行为、主线程阻塞的取舍——请先看[使用指南](/zh/packages/progress-overlay/guide/01-usage)。

## showProgressOverlay

```ts
showProgressOverlay(options?: ProgressOverlayOptions): ProgressOverlayHandle
```

创建遮罩（或加入并发任务已经在共用的那一个）并返回句柄。没有 `document` 时（Node/SSR），或 `container` 解析为 `null` 时，返回空操作句柄——同一处调用在两种环境下都能工作。

样式在**首次使用时**注入一次，即 `document.head` 里的 `<style id="mxe-overlay-style">`；之后的调用复用同一份。

## nextPaint

```ts
nextPaint(): Promise<void>
```

让出两个动画帧，并与 250ms 定时器赛跑——在隐藏标签页里 `requestAnimationFrame` 会被无限期暂停，定时器胜出，因此后台任务不会卡在让出上。这是在长同步段之前保证浏览器先画出遮罩的唯一手段；见[主线程阻塞](/zh/packages/progress-overlay/guide/01-usage#主线程阻塞)。

## ProgressOverlayOptions

| 配置项             | 类型                          | 默认值                                 | 说明                                                                                         |
| ------------------ | ----------------------------- | -------------------------------------- | -------------------------------------------------------------------------------------------- |
| `delayMs`          | `number`                      | `200`                                  | 挂载前延迟。任务在此期间结束则遮罩完全不出现。负值会被钳到 `0`                               |
| `minVisibleMs`     | `number`                      | `300`                                  | 一经显示至少保留这么久——用延时移除替代一闪而过。负值会被钳到 `0`                             |
| `fadeOutMs`        | `number`                      | `150`                                  | 节点摘除前的淡出时长。负值会被钳到 `0`                                                       |
| `zIndex`           | `number`                      | `2147483000`                           | 遮罩层级                                                                                     |
| `container`        | `HTMLElement`                 | `document.body`                        | 挂载目标                                                                                     |
| `blockInteraction` | `boolean`                     | `true`                                 | 在遮罩自身上拦截指针与滚动事件（绝不改宿主页面的 `overflow`）。置 `false` 时下层页面仍可操作 |
| `theme`            | `"auto" \| "light" \| "dark"` | `"auto"`                               | `"auto"` 在挂载时按 `prefers-color-scheme` 解析                                              |
| `text.title`       | `string`                      | `"请稍候"`                             | 面板标题                                                                                     |
| `text.initial`     | `string`                      | `"正在处理…"`                          | 收到首个 `setPhase` 前的文案                                                                 |
| `text.phases`      | `Record<string, string>`      | `{}`                                   | `key → 文案` 表；`setPhase(key)` 时查表。未覆盖的 key 原样显示                               |
| `text.hint`        | `string`                      | `"任务可能需要一些时间，请勿关闭页面"` | 文案下方的补充行，仅不确定态显示                                                             |

### 哪些配置项只对"建窗那次调用"生效

`container`、`theme`、`zIndex` 与 `blockInteraction` 属于**DOM 节点**的属性，而不属于某一次调用：只要并发任务还在共用同一个节点，它们就以**创建该节点的那次调用**为准。后到的调用若传了不同的值会被静默忽略，直到遮罩被完全拆除重建（发生在新调用携带了不同的 `container` 时）。文案项（`text.*`）是各调用方独立的，按[并发](/zh/packages/progress-overlay/guide/01-usage#并发)一节的规则合并进展示。

## ProgressOverlayHandle

| 方法             | 契约                                                                                                                                                                                                                          |
| ---------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `setProgress(p)` | 喂入进度。`(0, 1)` 之间的任何值都会把遮罩切到确定态进度条。开头那个 `0` 被忽略；收尾的 `1` 只会补完**已经处于确定态**的进度条——从未上报过中间进度的任务会一直停在旋转圆环上。进度值**没有单调性**，更小的值会把进度条画得更短 |
| `setPhase(key)`  | 按 `text.phases` 表切换文案。未覆盖的 key 原样显示，拼错会立刻暴露。重复传同一个 key 会被忽略（否则读屏会重复播报）                                                                                                           |
| `close()`        | 幂等，关闭后仍可安全调用。请放在 `finally` 里——不要由 `setProgress(1)` 触发关闭，因为失败路径也会收到它                                                                                                                       |

三个方法都做了异常安全：方法体整体被包裹，因为遮罩只是装饰，绝不允许弄挂它所装饰的任务。

接进度时还有两个行为值得知道：

- **揭示之前到达的值会被保留。** 若遮罩尚未揭示（还在 `delayMs` 窗口内），`setProgress` / `setPhase` 会更新状态但跳过渲染——揭示时会直接呈现这些值，而不是一个空面板。
- **句柄可能失效。** 若另一次调用把遮罩挂到了不同的 `container`，旧节点会被丢弃，指向它的句柄全部变成静默空操作（不抛错、不告警）。常见成因是并发任务之间混用了不同的 `container`——请统一。

## 外观定制

### CSS 变量

以下 9 个变量声明在 `.mxe-overlay` 上（浅色主题），并在 `.mxe-overlay[data-mxe-theme="dark"]` 下重新声明：

| 变量                 | 浅色                       | 深色                       |
| -------------------- | -------------------------- | -------------------------- |
| `--mxe-backdrop`     | `rgba(82, 82, 91, .32)`    | `rgba(0, 0, 0, .55)`       |
| `--mxe-panel-bg`     | `rgba(255, 255, 255, .72)` | `rgba(24, 24, 27, .65)`    |
| `--mxe-panel-border` | `rgba(0, 0, 0, .06)`       | `rgba(255, 255, 255, .08)` |
| `--mxe-panel-shadow` | 多层投影                   | 多层投影                   |
| `--mxe-title`        | `#18181b`                  | `#fafafa`                  |
| `--mxe-text`         | `#71717a`                  | `#a1a1aa`                  |
| `--mxe-hint`         | `#a1a1aa`                  | `#71717a`                  |
| `--mxe-track`        | `rgba(0, 0, 0, .08)`       | `rgba(255, 255, 255, .14)` |
| `--mxe-fill`         | `#171717`                  | `#fafafa`                  |

由于样式表是首次使用时**追加**到 `document.head` 的，它排在你应用启动时加载的样式表**之后**——所以同特异度的自定义样式会输。请提高特异度，而不要依赖顺序：

```css
/* 优先级高于注入的 `.mxe-overlay { --mxe-fill: #171717 }` */
html .mxe-overlay {
  --mxe-fill: #2563eb;
  --mxe-track: rgba(37, 99, 235, 0.18);
}
```

### 类名与数据属性

| 选择器 / 属性                                            | 元素                                                  |
| -------------------------------------------------------- | ----------------------------------------------------- |
| `.mxe-overlay`                                           | 根节点：背板、模糊、`aria-busy`，以及下面两个数据属性 |
| `.mxe-overlay[data-mxe-theme]`                           | `"light"` / `"dark"`——解析后的主题                    |
| `.mxe-overlay[data-mxe-mode]`                            | `"indeterminate"` / `"determinate"`——决定显示哪一半   |
| `.mxe-panel`                                             | 毛玻璃面板（标题、进度条行、提示）                    |
| `.mxe-spinner`、`.mxe-spinner-track`、`.mxe-spinner-arc` | 不确定态的 SVG 旋转圆环及其两个圆                     |
| `.mxe-title`                                             | 面板标题                                              |
| `.mxe-bar`、`.mxe-fill`                                  | 进度条轨道与填充                                      |
| `.mxe-row`、`.mxe-label`、`.mxe-percent`                 | 文案/百分比行；百分比节点在 `aria-live` 区域之外      |
| `.mxe-hint`                                              | 提示行（仅不确定态）                                  |

模式与主题正是两态行为的实现机制：`data-mxe-mode` 决定隐藏进度条与百分比（或隐藏旋转圆环与提示），因此自定义样式表可以围绕它们改外观，而不必触碰渲染逻辑。

## 谁在用它

`@marcusok/excel-exporter` 2.8 起通过它的 [`overlay` 选项](/zh/packages/excel-exporter/guide/11-overlay)替你驱动这套协议，并贡献导出语境的文案。当你需要给非导出任务加遮罩，或想自己驱动底层导出入口时，直接引入本包。
