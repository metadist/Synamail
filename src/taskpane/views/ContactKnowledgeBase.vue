<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import { go } from '@/taskpane/router'

const { t } = useI18n()
</script>

<template>
  <section class="ckb">
    <header class="syn-view-header">
      <button type="button" class="syn-back" @click="go('read')">← {{ t('common.back') }}</button>
      <h2 class="syn-view-title">{{ t('read.contactKb') }}</h2>
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
</style>
