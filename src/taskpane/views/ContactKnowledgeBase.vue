<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import { go } from '@/taskpane/router'

const { t } = useI18n()
</script>

<template>
  <section class="ckb">
    <header class="syn-row">
      <button type="button" class="ckb__back" @click="go('read')">← {{ t('common.back') }}</button>
      <h2>Contact knowledge base</h2>
    </header>
    <p class="syn-muted">
      {{ t('common.comingSoon') }} — per-contact RAG group search + "Save to contact" + "Ask about
      contact" lands in Sprint 3 (Step 3.5). See docs/FEATURES.md §4.
    </p>
  </section>
</template>

<style scoped>
.ckb {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.ckb__back {
  background: none;
  border: 0;
  color: var(--syn-brand-600);
  padding: 0;
}
</style>
