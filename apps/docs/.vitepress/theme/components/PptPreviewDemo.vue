<script setup lang="ts">
import { computed, onUnmounted, ref, shallowRef } from "vue";
import { useData } from "vitepress";

const { lang } = useData();
const isEn = computed(() => lang.value === "en-US");

const containerRef = ref<HTMLElement | null>(null);
const fileName = ref<string | null>(null);
const info = ref<{
  slideCount: number;
  firstSlideMs: number;
  durationMs: number;
} | null>(null);
const error = ref<{ code: string; message: string } | null>(null);
// 实例持有：createPptPreview 返回句柄，卸载/换文件时 destroy
const instance = shallowRef<{
  destroy(): void;
  nextSlide(): void;
  prevSlide(): void;
} | null>(null);
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
        title: "Drop a .pptx file",
        hint: "Parsing runs in a Web Worker (Rust/WASM); slides render lazily as they enter the viewport (far-away slides release their canvas). The text layer is selectable and copyable; slide navigation is 0-based. Password-protected (MS-OFFCRYPTO Agile) files decrypt when a password is provided.",
        choose: "Choose file",
        password: "Password (encrypted files)",
        slides: "slides",
        first: "first slide painted",
        total: "end-to-end",
        prev: "Prev",
        next: "Next",
        unsupported: "Unsupported file",
      }
    : {
        title: "拖入 .pptx 文件",
        hint: "解析在 Web Worker 内完成（Rust/WASM）；幻灯片进入视口才渲染（远离视口的页自动释放 canvas）。文本层可选中复制；页导航为 0-based。加密（MS-OFFCRYPTO Agile）文件提供密码即可解密。",
        choose: "选择文件",
        password: "密码（加密文件）",
        slides: "页",
        first: "首张出图",
        total: "端到端",
        prev: "上一页",
        next: "下一页",
        unsupported: "暂不支持的文件",
      },
);

const errorTexts: Record<string, string> = {
  UNSUPPORTED: isEn.value
    ? "Not a .pptx presentation (or an unsupported feature/resource budget exceeded)."
    : "不是 .pptx（或超出支持范围/资源预算）。",
  LEGACY_FORMAT: isEn.value
    ? "Legacy binary .ppt is not supported; re-save as .pptx."
    : "不支持老式 .ppt，请另存为 .pptx。",
  PASSWORD_PROTECTED: isEn.value
    ? "The file is encrypted: supply the correct password."
    : "文件已加密：请输入正确密码。",
  CORRUPT: isEn.value ? "The file is corrupted." : "文件结构损坏。",
  UNKNOWN: isEn.value ? "Preview failed" : "预览失败",
};

const password = ref("");

async function onFile(e: Event) {
  const input = e.target as HTMLInputElement;
  const file = input.files?.[0];
  const container = containerRef.value;
  if (!file || !container) return;
  const seq = ++loadSeq;
  info.value = null;
  error.value = null;
  fileName.value = file.name;
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const mod = await import("@marcusok/ppt-preview");
    // 等待期间又选了别的文件或组件已卸载：本次结果作废（旧预览保持不动）
    if (seq !== loadSeq) return;
    instance.value?.destroy();
    instance.value = null;
    container.textContent = "";
    instance.value = mod.createPptPreview(container, {
      source: bytes,
      ...(password.value.trim() !== ""
        ? { password: password.value.trim() }
        : {}),
      onRendered: (i) => {
        info.value = {
          slideCount: i.slideCount,
          firstSlideMs: i.firstSlideMs,
          durationMs: i.durationMs,
        };
      },
      onError: (err) => {
        error.value = { code: err.code, message: err.message };
      },
    });
  } catch (err) {
    if (seq !== loadSeq) return;
    error.value = {
      code: "UNKNOWN",
      message: err instanceof Error ? err.message : String(err),
    };
  }
}
</script>

<template>
  <div class="pptd-demo">
    <div class="pptd-toolbar">
      <label class="pptd-choose">
        <input type="file" accept=".pptx" @change="onFile" />
        {{ t.choose }}
      </label>
      <input
        v-model="password"
        class="pptd-pw"
        type="password"
        :placeholder="t.password"
        autocomplete="off"
      />
      <span v-if="fileName" class="pptd-name">{{ fileName }}</span>
    </div>
    <p class="pptd-hint">{{ t.hint }}</p>

    <p v-if="error" class="pptd-error">
      <strong>{{ errorTexts[error.code] ?? t.unsupported }}</strong>
      <span class="pptd-msg">{{ error.message }}</span>
    </p>

    <p v-if="info" class="pptd-info">
      <span>{{ info.slideCount }} {{ t.slides }}</span>
      <span>{{ t.first }} {{ info.firstSlideMs }}ms</span>
      <span>{{ t.total }} {{ info.durationMs }}ms</span>
      <button class="pptd-nav" @click="instance?.prevSlide()">
        {{ t.prev }}
      </button>
      <button class="pptd-nav" @click="instance?.nextSlide()">
        {{ t.next }}
      </button>
    </p>

    <div ref="containerRef" class="pptd-canvas"></div>
  </div>
</template>

<style scoped>
.pptd-demo {
  border: 1px solid var(--vp-c-border);
  border-radius: 8px;
  padding: 16px;
  margin: 16px 0;
}
.pptd-toolbar {
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
}
.pptd-choose {
  display: inline-block;
  border: 1px solid var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
  border-radius: 6px;
  padding: 4px 14px;
  cursor: pointer;
  font-size: 14px;
}
.pptd-choose input {
  display: none;
}
.pptd-pw {
  border: 1px solid var(--vp-c-border);
  border-radius: 6px;
  padding: 4px 10px;
  font-size: 13px;
  background: var(--vp-c-bg-alt);
  color: var(--vp-c-text-1);
}
.pptd-name {
  font-size: 14px;
  opacity: 0.75;
}
.pptd-hint {
  font-size: 13px;
  opacity: 0.65;
  margin: 8px 0 12px;
}
.pptd-error {
  border-radius: 6px;
  padding: 8px 12px;
  background: var(--vp-button-danger-bg);
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.pptd-msg {
  font-size: 12px;
  opacity: 0.8;
}
.pptd-info {
  display: flex;
  gap: 14px;
  flex-wrap: wrap;
  align-items: center;
  font-size: 13px;
  opacity: 0.85;
}
.pptd-nav {
  border: 1px solid var(--vp-c-border);
  border-radius: 6px;
  background: var(--vp-c-bg-alt);
  color: var(--vp-c-text-1);
  padding: 2px 10px;
  font-size: 13px;
  cursor: pointer;
}
.pptd-canvas {
  max-height: 480px;
  position: relative;
  border: 1px solid var(--vp-c-border);
  border-radius: 6px;
  overflow: auto;
  background: var(--vp-c-bg-alt);
}
</style>
