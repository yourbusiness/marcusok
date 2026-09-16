<script setup lang="ts">
import { computed } from "vue";
import { useData } from "vitepress";
import { useVisiblePackages } from "../composables/useVisiblePackages";

const { lang } = useData();
const isEn = computed(() => lang.value === "en-US");
const { visiblePackages } = useVisiblePackages();

// Line icon set (Lucide paths, 24x24 stroke=currentColor) rendered inside the
// icon tile. Registry entries reference icons by key; a key missing here
// falls back to rendering the raw string, so a stray value never shows an
// empty box. Static markup constants, safe for v-html.
const ICONS: Record<string, string> = {
  zap: '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>',
  pen: '<path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
  route:
    '<circle cx="6" cy="19" r="3"/><path d="M9 19h8.5a3.5 3.5 0 0 0 0-7h-11a3.5 3.5 0 0 1 0-7H15"/><circle cx="18" cy="5" r="3"/>',
  shield:
    '<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1 1 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/>',
};

// Every package contributes its own `highlights` from the registry, so the
// home page grows automatically when a new package is added.
const highlights = computed(() =>
  visiblePackages.value.flatMap((p) =>
    (p.highlights ?? []).map((h) => ({
      npmName: p.npmName,
      icon: h.icon,
      iconSvg: ICONS[h.icon],
      title: isEn.value ? h.title.en : h.title.zh,
      details: isEn.value ? h.details.en : h.details.zh,
    })),
  ),
);
</script>

<template>
  <div class="highlight-grid">
    <div
      v-for="(h, i) in highlights"
      :key="`${h.npmName}-${i}`"
      class="highlight-card"
    >
      <div class="highlight-card__icon">
        <svg
          v-if="h.iconSvg"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          stroke-width="2"
          stroke-linecap="round"
          stroke-linejoin="round"
          aria-hidden="true"
          v-html="h.iconSvg"
        />
        <template v-else>{{ h.icon }}</template>
      </div>
      <h3 class="highlight-card__title">{{ h.title }}</h3>
      <p class="highlight-card__details">{{ h.details }}</p>
    </div>
  </div>
</template>
