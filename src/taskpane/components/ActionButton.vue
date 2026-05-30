<script setup lang="ts">
interface Props {
  variant?: 'primary' | 'secondary' | 'danger'
  disabled?: boolean
  loading?: boolean
  block?: boolean
}
withDefaults(defineProps<Props>(), {
  variant: 'secondary',
  disabled: false,
  loading: false,
  block: true,
})
defineEmits<{ (e: 'click'): void }>()
</script>

<template>
  <button
    :class="['ab', `ab--${variant}`, { 'ab--block': block, 'ab--loading': loading }]"
    :disabled="disabled || loading"
    type="button"
    @click="$emit('click')"
  >
    <span v-if="loading" class="ab__spinner" aria-hidden="true" />
    <slot />
  </button>
</template>

<style scoped>
.ab {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: var(--syn-space-2);
  padding: var(--syn-space-2) var(--syn-space-3);
  border-radius: var(--syn-radius-md);
  border: 1px solid var(--syn-border);
  background: var(--syn-surface);
  color: var(--syn-text);
  font-weight: 500;
  min-height: 32px;
}
.ab--block {
  width: 100%;
}
.ab--primary {
  background: var(--syn-brand-600);
  border-color: var(--syn-brand-700);
  color: white;
}
.ab--primary:hover:not(:disabled) {
  background: var(--syn-brand-700);
}
.ab--danger {
  color: var(--syn-danger);
  border-color: var(--syn-danger);
  background: var(--syn-surface);
}
.ab--danger:hover:not(:disabled) {
  background: var(--syn-danger);
  color: white;
}
.ab:disabled {
  opacity: 0.55;
  cursor: not-allowed;
}
.ab__spinner {
  width: 12px;
  height: 12px;
  border-radius: 50%;
  border: 2px solid currentColor;
  border-top-color: transparent;
  animation: ab-spin 0.7s linear infinite;
}
@keyframes ab-spin {
  to {
    transform: rotate(360deg);
  }
}
</style>
