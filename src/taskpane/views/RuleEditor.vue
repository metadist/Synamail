<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import { go } from '@/taskpane/router'

const { t } = useI18n()
</script>

<template>
  <section class="rule">
    <header class="syn-row">
      <button type="button" class="rule__back" @click="go('settings')">
        ← {{ t('common.back') }}
      </button>
      <h2>{{ t('settings.routingRules') }}</h2>
    </header>
    <p class="syn-muted">
      {{ t('common.comingSoon') }} — Synapse Routing rules (BSELECTION_RULES) edit + preview lands
      in Sprint 3 (Step 3.7). See docs/FEATURES.md §5.
    </p>
  </section>
</template>

<style scoped>
.rule {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.rule__back {
  background: none;
  border: 0;
  color: var(--syn-brand-600);
  padding: 0;
}
</style>
