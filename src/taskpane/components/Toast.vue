<script setup lang="ts">
interface Props {
  kind?: 'info' | 'success' | 'error'
  message: string
}
const props = withDefaults(defineProps<Props>(), { kind: 'info' })

const glyph: Record<NonNullable<Props['kind']>, string> = {
  info: 'ℹ',
  success: '✓',
  error: '✕',
}
</script>

<template>
  <div :class="['toast', `toast--${props.kind}`]" role="status">
    <span class="toast__icon" aria-hidden="true">{{ glyph[props.kind] }}</span>
    <span class="toast__msg">{{ message }}</span>
  </div>
</template>

<style scoped>
.toast {
  display: flex;
  align-items: flex-start;
  gap: var(--syn-space-2);
  padding: var(--syn-space-2) var(--syn-space-3);
  border-radius: var(--syn-radius-md);
  border: 1px solid var(--syn-border);
  border-left-width: 3px;
  background: var(--syn-surface);
  color: var(--syn-text);
  font-size: var(--syn-font-size-sm);
  line-height: 1.35;
}
.toast__icon {
  font-weight: 700;
  line-height: 1.35;
  flex: none;
}
.toast__msg {
  min-width: 0;
  word-break: break-word;
}
.toast--success {
  background: var(--syn-success-bg);
  border-color: var(--syn-border);
  border-left-color: var(--syn-success);
  color: var(--syn-success-fg);
}
.toast--error {
  background: var(--syn-danger-bg);
  border-color: var(--syn-border);
  border-left-color: var(--syn-danger);
  color: var(--syn-danger-fg);
}
.toast--info {
  background: var(--syn-info-bg);
  border-color: var(--syn-border);
  border-left-color: var(--syn-brand-600);
  color: var(--syn-info-fg);
}
</style>
