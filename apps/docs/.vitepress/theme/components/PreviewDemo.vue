<script setup lang="ts">
import { computed, onUnmounted, ref, shallowRef } from "vue";
import { useData } from "vitepress";

const { lang } = useData();
const isEn = computed(() => lang.value === "en-US");

const containerRef = ref<HTMLElement | null>(null);
const fileName = ref<string | null>(null);
const info = ref<{
  sheetCount: number;
  rowCount: number;
  colCount: number;
  parseMs: number;
  renderMs: number;
  totalMs: number;
} | null>(null);
const error = ref<{ code: string; message: string } | null>(null);
const hasFile = ref(false);
const busy = ref(false);
// 实例持有：createPreview 返回句柄，卸载/换文件时 destroy
const instance = shallowRef<{ destroy(): void } | null>(null);

onUnmounted(() => {
  instance.value?.destroy();
  instance.value = null;
});

const t = computed(() =>
  isEn.value
    ? {
        title: "Drop an .xlsx / .xlsm / .csv file",
        hint: "Parsing runs in a Web Worker; rendering is DOM-virtualized (viewport cells only). Encrypted workbooks report a clear PASSWORD_PROTECTED error; legacy .xls gets a friendly unsupported message.",
        choose: "Choose file",
        sheets: "sheets",
        dims: "rows × cols",
        parse: "parse",
        render: "render",
        total: "total",
        unsupported: "Unsupported file",
      }
    : {
        title: "拖入 .xlsx / .xlsm / .csv 文件",
        hint: "解析在 Web Worker 内完成；渲染走 DOM 虚拟滚动（只渲染视口内格子）。加密文件会得到明确的 PASSWORD_PROTECTED 错误；旧版 .xls 会有友好的不支持提示。",
        choose: "选择文件",
        sheets: "个工作表",
        dims: "行 × 列",
        parse: "解析",
        render: "渲染",
        total: "合计",
        unsupported: "暂不支持的文件",
      },
);

const errorTexts: Record<string, string> = {
  PASSWORD_PROTECTED: isEn.value
    ? "Password-protected workbook (Agile AES-256) — pass the `password` option."
    : "工作簿已加密（Agile AES-256）——需要传入 `password` 选项。",
  LEGACY_FORMAT: isEn.value
    ? "Legacy .xls (BIFF8) is not supported; re-save as .xlsx."
    : "旧版 .xls（BIFF8）不支持，请另存为 .xlsx。",
  UNSUPPORTED: isEn.value ? "Unrecognized format." : "无法识别的文件格式。",
  CORRUPT: isEn.value
    ? "Not a valid xlsx (ZIP expected)."
    : "不是有效的 xlsx（需要 ZIP 结构）。",
  WASM: "WebAssembly unavailable.",
};

async function onFile(e: Event) {
  const input = e.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file || !containerRef.value) return;
  busy.value = true;
  info.value = null;
  error.value = null;
  instance.value?.destroy();
  instance.value = null;
  containerRef.value.textContent = "";
  fileName.value = file.name;
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    hasFile.value = true;
    const mod = await import("@marcusok/excel-preview");
    // 零配置：wasm/worker 资产默认随包定位（bundler 资产管线处理），
    // 与 ExportDemo 同一模式，无需 configureWasm。
    instance.value = mod.createPreview(containerRef.value, {
      source: bytes,
      onParsed: (i) => {
        info.value = {
          sheetCount: i.sheetCount,
          rowCount: i.rowCount,
          colCount: i.colCount,
          parseMs: i.duration.parse,
          renderMs: i.duration.render,
          totalMs: i.duration.total,
        };
      },
      onError: (err) => {
        error.value = { code: err.code, message: err.message };
      },
    });
  } catch (err) {
    error.value = {
      code: "UNKNOWN",
      message: err instanceof Error ? err.message : String(err),
    };
  } finally {
    busy.value = false;
  }
}
</script>

<template>
  <div class="pd-demo">
    <div class="pd-toolbar">
      <label class="pd-choose">
        <input type="file" accept=".xlsx,.xlsm,.csv" @change="onFile" />
        {{ t.choose }}
      </label>
      <span v-if="fileName" class="pd-name">{{ fileName }}</span>
    </div>
    <p class="pd-hint">{{ t.hint }}</p>

    <p v-if="error" class="pd-error">
      <strong>{{ errorTexts[error.code] ?? t.unsupported }}</strong>
      <span class="pd-msg">{{ error.message }}</span>
    </p>

    <p v-if="info" class="pd-info">
      <span>{{ info.sheetCount }} {{ t.sheets }}</span>
      <span
        >{{ info.rowCount.toLocaleString() }} × {{ info.colCount }}
        {{ t.dims }}</span
      >
      <span>{{ t.parse }} {{ info.parseMs }}ms</span>
      <span>{{ t.render }} {{ info.renderMs }}ms</span>
      <span>{{ t.total }} {{ info.totalMs }}ms</span>
    </p>

    <div
      ref="containerRef"
      class="pd-canvas"
      :class="{ 'pd-busy': busy }"
    ></div>
  </div>
</template>

<style scoped>
.pd-demo {
  border: 1px solid var(--vp-c-border);
  border-radius: 8px;
  padding: 16px;
  margin: 16px 0;
}
.pd-toolbar {
  display: flex;
  align-items: center;
  gap: 12px;
}
.pd-choose {
  display: inline-block;
  border: 1px solid var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
  border-radius: 6px;
  padding: 4px 14px;
  cursor: pointer;
  font-size: 14px;
}
.pd-choose input {
  display: none;
}
.pd-name {
  font-size: 14px;
  opacity: 0.75;
}
.pd-hint {
  font-size: 13px;
  opacity: 0.65;
  margin: 8px 0 12px;
}
.pd-error {
  border-radius: 6px;
  padding: 8px 12px;
  background: var(--vp-button-danger-bg);
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.pd-msg {
  font-size: 12px;
  opacity: 0.8;
}
.pd-info {
  display: flex;
  gap: 14px;
  flex-wrap: wrap;
  font-size: 13px;
  opacity: 0.85;
}
.pd-canvas {
  height: 420px;
  position: relative;
  border: 1px solid var(--vp-c-border);
  border-radius: 6px;
  overflow: hidden;
}
</style>
