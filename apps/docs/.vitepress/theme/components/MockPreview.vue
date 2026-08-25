<script setup lang="ts">
import { computed, ref } from "vue";
import { useData } from "vitepress";
import {
  getDataset,
  headerDepth,
  leafColumns,
  previewRows,
  type MockColumn,
} from "../../../src/demos/excel-exporter/datasets";

const props = withDefaults(
  defineProps<{ dataset: string; rows?: number; seed?: number }>(),
  { rows: 5, seed: 42 },
);

const { lang } = useData();
const isEn = computed(() => lang.value === "en-US");
const seedRef = ref(props.seed);
const ds = computed(() => getDataset(props.dataset));
const items = computed(() =>
  previewRows(props.dataset, props.rows, seedRef.value),
);

// 数据列 = 叶子列（分组列不产生数据列）。
const leaves = computed(() => leafColumns(ds.value.columns));

interface HeaderCell {
  key: string;
  label: string;
  rowspan: number;
  colspan: number;
}

/** 与 excel-exporter flattenColumnTree 相同的语义：分组列横向跨叶子列，叶子列纵向跨满剩余表头行。 */
const headerRows = computed<HeaderCell[][]>(() => {
  const depth = headerDepth(ds.value.columns);
  const rows: HeaderCell[][] = Array.from({ length: depth }, () => []);
  let groupSeq = 0;
  const label = (c: MockColumn) => (isEn.value ? c.header.en : c.header.zh);
  (function walk(cols: MockColumn[], level: number) {
    for (const c of cols) {
      if (c.children?.length) {
        rows[level]!.push({
          key: `g${groupSeq++}`,
          label: label(c),
          rowspan: 1,
          colspan: leafColumns(c.children).length,
        });
        walk(c.children, level + 1);
      } else {
        rows[level]!.push({
          key: c.key ?? `l${groupSeq++}`,
          label: label(c),
          rowspan: depth - level,
          colspan: 1,
        });
      }
    }
  })(ds.value.columns, 0);
  return rows;
});

function reshuffle() {
  seedRef.value = Math.floor(Math.random() * 2 ** 31);
}
</script>

<template>
  <div class="mock-preview">
    <div class="mock-preview__toolbar">
      <span>Mock · {{ items.length }} rows · seed {{ seedRef }}</span>
      <button type="button" @click="reshuffle">
        {{ isEn ? "Reshuffle" : "换一批" }}
      </button>
    </div>
    <table>
      <thead>
        <tr v-for="(cells, i) in headerRows" :key="i">
          <th
            v-for="c in cells"
            :key="c.key"
            :rowspan="c.rowspan > 1 ? c.rowspan : undefined"
            :colspan="c.colspan > 1 ? c.colspan : undefined"
          >
            {{ c.label }}
          </th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(row, i) in items" :key="i">
          <td v-for="c in leaves" :key="c.key">
            {{ String(row[c.key!] ?? "") }}
          </td>
        </tr>
      </tbody>
    </table>
  </div>
</template>
