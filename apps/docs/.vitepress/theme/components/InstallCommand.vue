<script setup lang="ts">
import { computed, onBeforeUnmount, ref } from "vue";
import { useData } from "vitepress";
import { useVisiblePackages } from "../composables/useVisiblePackages";

const { lang } = useData();
const isEn = computed(() => lang.value === "en-US");
const { visiblePackages } = useVisiblePackages();

// Primary package = first registry entry, so the install line follows the
// registry instead of hardcoding a name.
const command = computed(
  () =>
    `npm i ${visiblePackages.value[0]?.npmName ?? "@marcusok/excel-exporter"}`,
);

const copied = ref(false);
let timer: ReturnType<typeof setTimeout> | undefined;

async function copy(): Promise<void> {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(command.value);
    } else {
      // Fallback for non-secure contexts: transient textarea + execCommand.
      const ta = document.createElement("textarea");
      ta.value = command.value;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
    }
    copied.value = true;
    clearTimeout(timer);
    timer = setTimeout(() => {
      copied.value = false;
    }, 2000);
  } catch {
    // Clipboard blocked: leave the command selected so it can be copied
    // manually; do not flip into the "copied" state.
  }
}

onBeforeUnmount(() => {
  clearTimeout(timer);
});
</script>

<template>
  <div class="install-command">
    <code class="install-command__text"
      ><span class="install-command__dollar">$</span> {{ command }}</code
    >
    <button
      class="install-command__copy"
      type="button"
      :aria-label="isEn ? 'Copy install command' : '复制安装命令'"
      @click="copy"
    >
      <svg
        v-if="!copied"
        viewBox="0 0 24 24"
        width="13"
        height="13"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
      >
        <rect width="14" height="14" x="8" y="8" rx="2" ry="2" />
        <path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2" />
      </svg>
      <svg
        v-else
        viewBox="0 0 24 24"
        width="13"
        height="13"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
      >
        <path d="M20 6 9 17l-5-5" />
      </svg>
      <span>{{
        copied ? (isEn ? "Copied" : "已复制") : isEn ? "Copy" : "复制"
      }}</span>
    </button>
  </div>
</template>
