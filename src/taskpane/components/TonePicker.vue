<script setup lang="ts">
import { useI18n } from 'vue-i18n'

const TONES = ['formal', 'concise', 'friendly'] as const
type Tone = (typeof TONES)[number]

defineProps<{ modelValue: Tone }>()
defineEmits<{ (e: 'update:modelValue', v: Tone): void }>()

const { t } = useI18n()
</script>

<template>
  <select
    class="tp"
    :value="modelValue"
    @change="$emit('update:modelValue', ($event.target as HTMLSelectElement).value as Tone)"
  >
    <option v-for="tone in TONES" :key="tone" :value="tone">
      {{ t(`tone.${tone}`) }}
    </option>
  </select>
</template>

<style scoped>
.tp {
  padding: var(--syn-space-1) var(--syn-space-2);
  border-radius: var(--syn-radius-sm);
  border: 1px solid var(--syn-border);
  background: var(--syn-bg);
  color: var(--syn-text);
  font-size: var(--syn-font-size-sm);
}
</style>
