<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import { go } from '@/taskpane/router'

const { t } = useI18n()
</script>

<template>
  <section class="rule">
    <header class="syn-view-header">
      <button type="button" class="syn-back" @click="go('settings')">
        ← {{ t('common.back') }}
      </button>
      <h2 class="syn-view-title">{{ t('settings.routingRules') }}</h2>
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
</style>
