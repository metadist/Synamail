<script setup lang="ts">
defineProps<{ size?: number }>()
</script>

<template>
  <span
    class="sp"
    :style="{ width: `${size ?? 16}px`, height: `${size ?? 16}px` }"
    aria-hidden="true"
  />
</template>

<style scoped>
.sp {
  display: inline-block;
  border-radius: 50%;
  border: 2px solid var(--syn-brand-500);
  border-top-color: transparent;
  animation: sp 0.7s linear infinite;
}
@keyframes sp {
  to {
    transform: rotate(360deg);
  }
}
</style>
