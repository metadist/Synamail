<script setup lang="ts">
interface Props {
  kind?: 'info' | 'success' | 'error'
  message: string
}
withDefaults(defineProps<Props>(), { kind: 'info' })
</script>

<template>
  <div :class="['toast', `toast--${kind}`]" role="status">
    {{ message }}
  </div>
</template>

<style scoped>
.toast {
  padding: var(--syn-space-2) var(--syn-space-3);
  border-radius: var(--syn-radius-md);
  border: 1px solid var(--syn-border);
  background: var(--syn-surface);
  font-size: var(--syn-font-size-sm);
}
.toast--error {
  color: var(--syn-danger);
  border-color: rgba(185, 28, 28, 0.3);
}
.toast--success {
  color: var(--syn-success);
  border-color: rgba(21, 128, 61, 0.3);
}
</style>
