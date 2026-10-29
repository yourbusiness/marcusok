<script setup lang="ts">
import { computed, onUnmounted, ref, shallowRef } from "vue";
import { useData } from "vitepress";

const { lang } = useData();
const isEn = computed(() => lang.value === "en-US");

const containerRef = ref<HTMLElement | null>(null);
const fileName = ref<string | null>(null);
const info = ref<{ pageCount: number; durationMs: number } | null>(null);
const error = ref<{ code: string; message: string } | null>(null);
const busy = ref(false);
// 实例持有：createWordPreview 返回句柄，卸载/换文件时 destroy
const instance = shallowRef<{ destroy(): void } | null>(null);
// 快速连续选文件时，后一次加载必须作废前一次；否则两个实例可能挂到同一容器
let loadSeq = 0;

onUnmounted(() => {
  loadSeq++; // 作废仍在等待中的加载
  instance.value?.destroy();
  instance.value = null;
});

const t = computed(() =>
  isEn.value
    ? {
        title: "Drop a .docx / .docm file",
        hint: "Rendering runs on the main thread (engine constraint — no worker mode). Legacy .doc and encrypted files report clear error codes; hyperlinks outside the http/https/mailto/tel allowlist are neutralized.",
        choose: "Choose file",
        pages: "pages (sections)",
        rendered: "full render (incl. images/fonts)",
        unsupported: "Unsupported file",
      }
    : {
        title: "拖入 .docx / .docm 文件",
        hint: "渲染在主线程完成（引擎限制，无 Worker 模式）。.doc 老格式与加密文件会给出明确错误码；http/https/mailto/tel 白名单之外的超链接会被置空。",
        choose: "选择文件",
        pages: "页（section）",
        rendered: "完整渲染（含图片/字体）",
        unsupported: "暂不支持的文件",
      },
);

const errorTexts: Record<string, string> = {
  // OLE2 魔数无法区分 .doc 老格式与标准加密 docx，首版两者均不支持
  LEGACY_FORMAT: isEn.value
    ? "Legacy .doc and encrypted .docx files are both unsupported; re-save as a plain .docx."
    : ".doc 老格式与加密的 .docx 均不支持，请另存为普通 .docx。",
  PASSWORD_PROTECTED: isEn.value
    ? "Password-protected documents are not supported."
    : "不支持加密文档。",
  UNSUPPORTED: isEn.value
    ? "Not a Word document (missing required parts)."
    : "不是 Word 文档（缺少必需部件）。",
  CORRUPT: isEn.value
    ? "The file is corrupted or failed to parse."
    : "文件结构损坏或解析失败。",
};

async function onFile(e: Event) {
  const input = e.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file) return;
  const seq = ++loadSeq;
  busy.value = true;
  info.value = null;
  error.value = null;
  fileName.value = file.name;
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const mod = await import("@marcusok/word-preview");
    // 等待期间又选了别的文件或组件已卸载：本次结果作废（旧预览保持不动）
    if (seq !== loadSeq || !containerRef.value) return;
    instance.value?.destroy();
    instance.value = null;
    containerRef.value.textContent = "";
    instance.value = mod.createWordPreview(containerRef.value, {
      source: bytes,
      onRendered: (i) => {
        info.value = { pageCount: i.pageCount, durationMs: i.durationMs };
        busy.value = false;
      },
      onError: (err) => {
        error.value = { code: err.code, message: err.message };
        busy.value = false;
      },
    });
  } catch (err) {
    if (seq !== loadSeq) return;
    error.value = {
      code: "UNKNOWN",
      message: err instanceof Error ? err.message : String(err),
    };
    busy.value = false;
  }
}
</script>

<template>
  <div class="wpd-demo">
    <div class="wpd-toolbar">
      <label class="wpd-choose">
        <input type="file" accept=".docx,.docm" @change="onFile" />
        {{ t.choose }}
      </label>
      <span v-if="fileName" class="wpd-name">{{ fileName }}</span>
    </div>
    <p class="wpd-hint">{{ t.hint }}</p>

    <p v-if="error" class="wpd-error">
      <strong>{{ errorTexts[error.code] ?? t.unsupported }}</strong>
      <span class="wpd-msg">{{ error.message }}</span>
    </p>

    <p v-if="info" class="wpd-info">
      <span>{{ info.pageCount }} {{ t.pages }}</span>
      <span>{{ t.rendered }} {{ info.durationMs }}ms</span>
    </p>

    <div
      ref="containerRef"
      class="wpd-canvas"
      :class="{ 'wpd-busy': busy }"
    ></div>
  </div>
</template>

<style scoped>
.wpd-demo {
  border: 1px solid var(--vp-c-border);
  border-radius: 8px;
  padding: 16px;
  margin: 16px 0;
}
.wpd-toolbar {
  display: flex;
  align-items: center;
  gap: 12px;
}
.wpd-choose {
  display: inline-block;
  border: 1px solid var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
  border-radius: 6px;
  padding: 4px 14px;
  cursor: pointer;
  font-size: 14px;
}
.wpd-choose input {
  display: none;
}
.wpd-name {
  font-size: 14px;
  opacity: 0.75;
}
.wpd-hint {
  font-size: 13px;
  opacity: 0.65;
  margin: 8px 0 12px;
}
.wpd-error {
  border-radius: 6px;
  padding: 8px 12px;
  background: var(--vp-button-danger-bg);
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.wpd-msg {
  font-size: 12px;
  opacity: 0.8;
}
.wpd-info {
  display: flex;
  gap: 14px;
  flex-wrap: wrap;
  font-size: 13px;
  opacity: 0.85;
}
.wpd-canvas {
  max-height: 480px;
  position: relative;
  border: 1px solid var(--vp-c-border);
  border-radius: 6px;
  overflow: auto;
}
</style>
