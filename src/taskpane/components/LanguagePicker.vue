<script setup lang="ts">
import { useI18n } from 'vue-i18n'

const LANGS = ['auto', 'en', 'de', 'fr', 'es', 'it', 'zh', 'ar'] as const
type Lang = (typeof LANGS)[number]

defineProps<{ modelValue: Lang }>()
defineEmits<{ (e: 'update:modelValue', v: Lang): void }>()

const { t } = useI18n()
</script>

<template>
  <select
    class="lp"
    :value="modelValue"
    @change="$emit('update:modelValue', ($event.target as HTMLSelectElement).value as Lang)"
  >
    <option v-for="l in LANGS" :key="l" :value="l">
      {{ t(`language.${l}`) }}
    </option>
  </select>
</template>

<style scoped>
.lp {
  padding: var(--syn-space-1) var(--syn-space-2);
  border-radius: var(--syn-radius-sm);
  border: 1px solid var(--syn-border);
  background: var(--syn-bg);
  color: var(--syn-text);
  font-size: var(--syn-font-size-sm);
}
</style>
