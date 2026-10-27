<script setup lang="ts">
import { computed, onUnmounted, ref, shallowRef } from "vue";
import { useData, withBase } from "vitepress";

const { lang } = useData();
const isEn = computed(() => lang.value === "en-US");

const containerRef = ref<HTMLElement | null>(null);
const fileName = ref<string | null>(null);
const info = ref<{
  pageCount: number;
  firstPageMs: number;
  durationMs: number;
} | null>(null);
const error = ref<{ code: string; message: string } | null>(null);
// 实例持有：createPdfPreview 返回句柄，卸载/换文件时 destroy
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
        title: "Drop a .pdf file",
        hint: "Parsing runs in a Web Worker; pages render lazily as they enter the viewport (far-away pages release their canvas). The text layer is selectable and copyable. Encrypted PDFs report a clear error code; embedded PDF JavaScript never executes.",
        choose: "Choose file",
        pages: "pages",
        first: "first page painted",
        total: "end-to-end",
        unsupported: "Unsupported file",
      }
    : {
        title: "拖入 .pdf 文件",
        hint: "解析在 Web Worker 内完成；页面进入视口才渲染（远离视口的页自动释放 canvas）。文本层可选中复制。加密 PDF 给出明确错误码；内嵌 PDF JavaScript 永不执行。",
        choose: "选择文件",
        pages: "页",
        first: "首页出图",
        total: "端到端",
        unsupported: "暂不支持的文件",
      },
);

const errorTexts: Record<string, string> = {
  UNSUPPORTED: isEn.value
    ? "Not a PDF (no %PDF- magic within the first 1024 bytes)."
    : "不是 PDF（头部 1024 字节内无 %PDF- 魔数）。",
  PASSWORD_PROTECTED: isEn.value
    ? "Password-protected PDFs are not supported."
    : "不支持加密 PDF。",
  CORRUPT: isEn.value
    ? "The file is corrupted or failed to parse."
    : "文件结构损坏或解析失败。",
  UNKNOWN: isEn.value ? "Preview failed" : "预览失败",
};

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
    const mod = await import("@marcusok/pdf-preview");
    // 等待期间又选了别的文件或组件已卸载：本次结果作废（旧预览保持不动）
    if (seq !== loadSeq) return;
    instance.value?.destroy();
    instance.value = null;
    container.textContent = "";
    // worker 零配置（bundler 对锚点字面量自动 emit、URL 自动 base-aware）；
    // cmaps/standard_fonts 是目录资产，bundler 不会复制——runtimeAssets 已
    // 把它们拷入 public/，这里显式传带 base 的 URL。注意 withBase 只对 "/"
    // 开头的路径应用 base（不带斜杠会原样返回相对路径，被 pdf.js 相对页面
    // 解析成 /marcusok/packages/pdf-preview/cmaps/ → 404、中文静默劣化——
    // 浏览器验收实测踩过，必须带前导斜杠）
    instance.value = mod.createPdfPreview(container, {
      source: bytes,
      cMapUrl: withBase("/cmaps/"),
      standardFontDataUrl: withBase("/standard_fonts/"),
      onRendered: (i) => {
        info.value = {
          pageCount: i.pageCount,
          firstPageMs: i.firstPageMs,
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
  <div class="ppd-demo">
    <div class="ppd-toolbar">
      <label class="ppd-choose">
        <input type="file" accept=".pdf" @change="onFile" />
        {{ t.choose }}
      </label>
      <span v-if="fileName" class="ppd-name">{{ fileName }}</span>
    </div>
    <p class="ppd-hint">{{ t.hint }}</p>

    <p v-if="error" class="ppd-error">
      <strong>{{ errorTexts[error.code] ?? t.unsupported }}</strong>
      <span class="ppd-msg">{{ error.message }}</span>
    </p>

    <p v-if="info" class="ppd-info">
      <span>{{ info.pageCount }} {{ t.pages }}</span>
      <span>{{ t.first }} {{ info.firstPageMs }}ms</span>
      <span>{{ t.total }} {{ info.durationMs }}ms</span>
    </p>

    <div ref="containerRef" class="ppd-canvas"></div>
  </div>
</template>

<style scoped>
.ppd-demo {
  border: 1px solid var(--vp-c-border);
  border-radius: 8px;
  padding: 16px;
  margin: 16px 0;
}
.ppd-toolbar {
  display: flex;
  align-items: center;
  gap: 12px;
}
.ppd-choose {
  display: inline-block;
  border: 1px solid var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
  border-radius: 6px;
  padding: 4px 14px;
  cursor: pointer;
  font-size: 14px;
}
.ppd-choose input {
  display: none;
}
.ppd-name {
  font-size: 14px;
  opacity: 0.75;
}
.ppd-hint {
  font-size: 13px;
  opacity: 0.65;
  margin: 8px 0 12px;
}
.ppd-error {
  border-radius: 6px;
  padding: 8px 12px;
  background: var(--vp-button-danger-bg);
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.ppd-msg {
  font-size: 12px;
  opacity: 0.8;
}
.ppd-info {
  display: flex;
  gap: 14px;
  flex-wrap: wrap;
  font-size: 13px;
  opacity: 0.85;
}
.ppd-canvas {
  max-height: 480px;
  position: relative;
  border: 1px solid var(--vp-c-border);
  border-radius: 6px;
  overflow: auto;
  background: var(--vp-c-bg-alt);
}
</style>
