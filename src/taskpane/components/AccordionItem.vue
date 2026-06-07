<script setup lang="ts">
/**
 * A single accordion "flap" built on the native <details>/<summary> element —
 * accessible and keyboard-friendly for free. `open` only seeds the initial
 * state; the user toggles freely afterwards. Content is always in the DOM
 * (just hidden when collapsed), which keeps embedded forms' state alive across
 * open/close.
 */
interface Props {
  title: string
  subtitle?: string
  open?: boolean
  /** Render the subtitle in full text color + bold (e.g. an email subject). */
  strongSubtitle?: boolean
}
withDefaults(defineProps<Props>(), { subtitle: '', open: false, strongSubtitle: false })
</script>

<template>
  <details class="acc" :open="open">
    <summary class="acc__summary">
      <span class="acc__heading">
        <span class="acc__title">{{ title }}</span>
        <span v-if="subtitle" class="acc__sub" :class="{ 'acc__sub--strong': strongSubtitle }">{{
          subtitle
        }}</span>
      </span>
      <span class="acc__chev" aria-hidden="true">▾</span>
    </summary>
    <div class="acc__body">
      <slot />
    </div>
  </details>
</template>

<style scoped>
.acc {
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-md);
  background: var(--syn-surface);
  overflow: hidden;
}
.acc__summary {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--syn-space-2);
  padding: var(--syn-space-3);
  cursor: pointer;
  list-style: none;
  user-select: none;
}
/* Hide the default disclosure triangle across engines. */
.acc__summary::-webkit-details-marker {
  display: none;
}
.acc__summary:hover {
  background: var(--syn-bg);
}
.acc__heading {
  display: flex;
  flex-direction: column;
  gap: 2px;
  min-width: 0;
}
.acc__title {
  font-weight: 600;
  color: var(--syn-text);
}
.acc__sub {
  font-size: var(--syn-font-size-sm);
  color: var(--syn-muted);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.acc__sub--strong {
  color: var(--syn-text);
  font-weight: 600;
}
.acc__chev {
  color: var(--syn-muted);
  transition: transform 0.15s ease;
  flex: none;
}
.acc[open] .acc__chev {
  transform: rotate(180deg);
}
.acc__body {
  padding: 0 var(--syn-space-3) var(--syn-space-3);
  border-top: 1px solid var(--syn-border);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
  padding-top: var(--syn-space-3);
}
</style>
