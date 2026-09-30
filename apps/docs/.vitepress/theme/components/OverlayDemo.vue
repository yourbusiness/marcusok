<script setup lang="ts">
import { computed, onUnmounted, ref } from "vue";
import { useData } from "vitepress";
import type {
  ProgressOverlayHandle,
  ProgressOverlayOptions,
} from "@marcusok/progress-overlay";

/**
 * Live demo for @marcusok/progress-overlay (registry: demo: "OverlayDemo").
 *
 * 演示刻意不使用任何业务词汇：本包与 Excel 无关，文案全部由调用方给出
 * （包内默认文案是中文的"请稍候"，英文站必须自带 text 覆盖）。
 *
 * 任务分两段，正是两个状态的对照：
 *   前 1.5s 无进度 → 不确定态（旋转圆环 + hint）
 *   随后流式推进   → 确定态（百分比条）
 */
const { lang } = useData();
const isEn = computed(() => lang.value === "en-US");

type ThemeChoice = NonNullable<ProgressOverlayOptions["theme"]>;
const themes: ThemeChoice[] = ["auto", "light", "dark"];

const theme = ref<ThemeChoice>("auto");
const delayed = ref(true);
const running = ref(false);
const lastRun = ref<string | null>(null);

// 卸载兜底：SPA 导航离开后任务可能仍在跑，句柄持有的遮罩会留在页面上；
// 同时置 null 避免回调触碰已卸载组件的响应式状态。
const handle = ref<ProgressOverlayHandle | null>(null);
const disposed = ref(false);
onUnmounted(() => {
  disposed.value = true;
  handle.value?.close();
  handle.value = null;
});

const t = computed(() =>
  isEn.value
    ? {
        theme: "Theme",
        delay: "delayMs 200 (short tasks never flash)",
        run: "Run a 3.5s task",
        running: "Running…",
        done: (ms: number) =>
          `Task finished in ${ms}ms — the overlay closed itself.`,
        hint: "The task reports no progress for its first 1.5s (spinner + hint), then streams progress (percentage bar). Both states render into one glass panel; the caller only needs setProgress() / setPhase() / close().",
        text: {
          title: "Syncing data",
          initial: "Connecting…",
          phases: {
            pulling: "Pulling remote data…",
            saving: "Saving locally…",
          },
          hint: "This demo task runs for ~3.5s",
        },
      }
    : {
        theme: "主题",
        delay: "delayMs 200（短任务不闪现）",
        run: "运行 3.5 秒任务",
        running: "运行中…",
        done: (ms: number) => `任务完成，总耗时 ${ms}ms，遮罩已自动关闭。`,
        hint: "任务前 1.5 秒不上报进度（旋转圆环 + 提示行），随后流式推进（百分比条）。两种状态渲染在同一个毛玻璃面板里，调用方只需要 setProgress() / setPhase() / close()。",
        text: {
          title: "正在同步数据",
          initial: "正在连接…",
          phases: {
            pulling: "正在拉取远端数据…",
            saving: "正在写入本地…",
          },
          hint: "本演示任务约 3.5 秒",
        },
      },
);

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function run() {
  if (running.value) return;
  running.value = true;
  lastRun.value = null;
  const startedAt = performance.now();
  try {
    // 动态 import：与 ExportDemo / PreviewDemo 同一模式，遮罩包只在首次
    // 点击演示时加载。
    const { showProgressOverlay } = await import("@marcusok/progress-overlay");
    const overlay = showProgressOverlay({
      delayMs: delayed.value ? 200 : 0,
      theme: theme.value,
      text: t.value.text,
    });
    handle.value = overlay;
    try {
      await sleep(1500);
      overlay.setPhase("pulling");
      for (const p of [0.12, 0.3, 0.5, 0.72, 0.9]) {
        await sleep(280);
        overlay.setProgress(p);
      }
      overlay.setPhase("saving");
      await sleep(200);
      overlay.setProgress(1);
      if (!disposed.value) {
        lastRun.value = t.value.done(Math.round(performance.now() - startedAt));
      }
    } finally {
      // close 必须放 finally：它同时承担失败路径的收尾。
      overlay.close();
      handle.value = null;
    }
  } finally {
    if (!disposed.value) running.value = false;
  }
}
</script>

<template>
  <div class="demo-panel">
    <div class="demo-panel__controls">
      <span class="demo-panel__label">{{ t.theme }}</span>
      <button
        v-for="item in themes"
        :key="item"
        type="button"
        class="demo-panel__theme"
        :class="{ 'is-active': theme === item }"
        :disabled="running"
        @click="theme = item"
      >
        {{ item }}
      </button>
      <label>
        <input v-model="delayed" type="checkbox" :disabled="running" />
        {{ t.delay }}
      </label>
      <button type="button" :disabled="running" @click="run">
        {{ running ? t.running : t.run }}
      </button>
    </div>

    <div v-if="lastRun" class="demo-panel__result">{{ lastRun }}</div>
    <p v-else class="demo-panel__note">{{ t.hint }}</p>
  </div>
</template>
